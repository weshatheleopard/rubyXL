module RubyXL
module LegacyWorksheet
  include Enumerable

  def initialize(params = {})
    super
    self.workbook   = params[:workbook]
    self.sheet_name = params[:sheet_name]
    self.sheet_id   = params[:sheet_id]
    self.sheet_data = RubyXL::SheetData.new
    self.cols = RubyXL::ColumnRanges.new
    @comments = [] # Do not optimize! These are arrays, so they will share the pointer!
    @printer_settings = []
    @generic_storage = []
  end

  # allows for easier access to sheet_data
  def [](row = 0)
    sheet_data[row]
  end

  def each
    sheet_data.rows.each { |row| yield(row) }
  end

  #returns 2d array of just the cell values (without style or formula information)
  def extract_data(args = {})
    sheet_data.rows.map { |row|
      row.cells.map { |c| c && c.value(args) } unless row.nil?
    }
  end

  def get_table(headers = [], opts = {})
    validate_workbook

    headers = [headers] unless headers.is_a?(Array)
    normalized_headers = headers.map {|v| normalize_value(v)}

    row_num = find_first_row_with_content(normalized_headers)
    return nil if row_num.nil?

    # find the header index of the last header
    last_header_i = if opts.has_key?(:last_header)
                      last_header = normalize_value(opts[:last_header])
                      normalized_headers.find_index { |hv| hv == last_header }
                    end

    # build a map of header index to column index
    hi2ci = {}
    sheet_data[row_num].cells.each_with_index do |cell, cell_i|
      cv = normalize_value(cell && cell.value)
      header_i = normalized_headers.find_index{|hv| hv == cv}
      hi2ci[header_i] = cell_i unless header_i.nil?
      break if header_i && header_i == last_header_i
    end

    # table_hash has a key for each header. Each key points to a separate array that will contain column values
    table_hash = Hash[hi2ci.keys.map{|hi| [headers[hi], []]}]

    # build sorted header list... sort headers by their corresponding cell index
    table_hash[:sorted_headers] = hi2ci.to_a.sort_by{|pair| pair[1]}.map{|pair| headers[pair[0]]} unless hi2ci.empty?

    # generate table_hash from each non-empty row of the table
    table_hash[:table] = (row_num + 1..sheet_data.rows.count).map do |i|
      row = sheet_data[i]
      next if row.nil?

      # collect cell values from row under each column header
      values = hi2ci.map {|hi, ci| c = row.cells[ci]; [headers[hi], (c.value unless c.nil?)]}

      # skip rows that consist of only empty cells
      next if values.all? {|v| v[1].nil?}

      # add each cell value to table_hash[header] columns
      values.each {|e| hv, cv = *e; table_hash[hv] << cv}

      # convert non-empty cell values to a hash and append to table_hash[:table]
      Hash[values.reject{|v| v[1].nil?}]
    end.compact

    table_hash
  end

  def normalize_value(v) v.to_s.strip.downcase end
  private :normalize_value

  # finds first row which contains at least all strings in cells_content
  def find_first_row_with_content(header_values)
    sheet_data.rows.find_index do |row|
      next if row.nil?
      row_values = row.cells.map {|cell| normalize_value(cell && cell.value) }

      (header_values & row_values).size == header_values.size
    end
  end
  private :find_first_row_with_content

  # Set raw column width value
  def change_column_width_raw(column_index, width)
    validate_workbook
    ensure_cell_exists(0, column_index)
    range = cols.get_range(column_index)
    range.width = width
    range.custom_width = true
  end

  # Get column width measured in number of digits, as per
  # http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column%28v=office.14%29.aspx
  def change_column_width(column_index, width_in_chars = RubyXL::ColumnRange::DEFAULT_WIDTH)
    change_column_width_raw(column_index, ((width_in_chars + (5.0 / RubyXL::Font::MAX_DIGIT_WIDTH)) * 256).to_i / 256.0)
  end

  def change_column_fill(column_index, color_index='ffffff')
    validate_workbook
    Color.validate_color(color_index)
    ensure_cell_exists(0, column_index)

    cols.get_range(column_index).style_index = @workbook.modify_fill(get_col_style(column_index), color_index)

    sheet_data.rows.each { |row|
      c = row[column_index]
      c.change_fill(color_index) if c
    }
  end

  # merges cells within a rectangular range
  def merge_cells(row1 = 0, col1 = 0, row2 = 0, col2 = 0)
    validate_workbook

    self.merged_cells ||= RubyXL::MergedCells.new
    merged_cells << RubyXL::MergedCell.new(:ref => RubyXL::Reference.new(row1, row2, col1, col2))
  end

  def add_row(row_index = 0, params = {})
    new_row = RubyXL::Row.new(params)
    new_row.worksheet = self
    sheet_data.rows[row_index] = new_row
  end

  def add_cell(row_index = 0, column_index = 0, data = '', formula = nil, overwrite = true)
    validate_workbook
    validate_nonnegative(row_index)
    validate_nonnegative(column_index)
    row = sheet_data.rows[row_index] || add_row(row_index)

    c = row.cells[column_index]

    if overwrite || c.nil?
      c = RubyXL::Cell.new
      c.worksheet = self
      c.row = row_index
      c.column = column_index
      c.raw_value = data
      c.datatype = RubyXL::DataType::RAW_STRING unless formula || data.is_a?(Numeric)
      c.formula = RubyXL::Formula.new(:expression => formula) if formula

      range = cols && cols.locate_range(column_index)
      c.style_index = row.style_index || (range && range.style_index) || 0
      row.cells[column_index] = c
    end

    c
  end

  def delete_row(row_index=0)
    validate_workbook
    validate_nonnegative(row_index)

    deleted = sheet_data.rows.delete_at(row_index)

    # Change cell row numbers
    row_index.upto(sheet_data.size - 1) { |index|
      sheet_data[index].cells.each{ |c| c.row -= 1 unless c.nil? }
    }

    return deleted
  end

  # Inserts row at row_index, pushes down, copies style from the row above (that's what Excel 2013 does!)
  # NOTE: use of this method will break formulas which reference cells which are being "pushed down"
  def insert_row(row_index = 0)
    validate_workbook
    ensure_cell_exists(row_index)

    old_row = new_cells = nil

    if row_index > 0 then
      old_row = sheet_data.rows[row_index - 1]
      if old_row then
        new_cells = old_row.cells.collect { |c|
                                            if c.nil? then nil
                                            else RubyXL::Cell.new(:style_index => c.style_index)
                                            end }
      end
    end

    row0 = sheet_data.rows[0]
    new_cells ||= Array.new((row0 && row0.cells.size) || 0)

    sheet_data.rows.insert(row_index, nil)
    new_row = add_row(row_index, :cells => new_cells, :style_index => old_row && old_row.style_index)

    # Update row values for all rows below
    row_index.upto(sheet_data.rows.size - 1) { |i|
      row = sheet_data.rows[i]
      next if row.nil?
      row.cells.each { |c| c.row = i unless c.nil? }
    }

    return new_row
  end

  def delete_column(column_index = 0)
    validate_workbook
    validate_nonnegative(column_index)

    #delete column
    sheet_data.rows.each { |row| row.cells.delete_at(column_index) }

    # Change column numbers for cells to the right of the deleted column
    sheet_data.rows.each_with_index { |row, row_index|
      row.cells.each_with_index { |c, column_index|
        c.column = column_index if c.is_a?(Cell)
      }
    }

    cols.each { |range| range.delete_column(column_index) }
  end

  # Inserts column at +column_index+, pushes everything right, takes styles from column to left
  # NOTE: use of this method will break formulas which reference cells which are being "pushed right"
  def insert_column(column_index = 0)
    validate_workbook
    ensure_cell_exists(0, column_index)

    old_range = cols.get_range(column_index)

    #go through each cell in column
    sheet_data.rows.each_with_index { |row, row_index|
      old_cell = row[column_index]
      c = nil

      if old_cell && old_cell.style_index != 0 &&
           old_range && old_range.style_index != old_cell.style_index then

        c = RubyXL::Cell.new(:style_index => old_cell.style_index, :worksheet => self,
                             :row => row_index, :column => column_index,
                             :datatype => RubyXL::DataType::SHARED_STRING)
      end

      row.insert_cell_shift_right(c, column_index)
    }

    cols.insert_column(column_index)

    # TODO: update column numbers
  end

  def insert_cell(row = 0, col = 0, data = nil, formula = nil, shift = nil)
    validate_workbook
    ensure_cell_exists(row, col)

    case shift
    when nil then # No shifting at all
    when :right then
      sheet_data.rows[row].insert_cell_shift_right(nil, col)
    when :down then
      add_row(sheet_data.size, :cells => Array.new(sheet_data.rows[row].size))
      (sheet_data.size - 1).downto(row+1) { |index|
        sheet_data.rows[index].cells[col] = sheet_data.rows[index-1].cells[col]
      }                                                                             		
    else
      raise 'invalid shift option'
    end

    return add_cell(row,col,data,formula)
  end

  # by default, only sets cell to nil
  # if :left is specified, method will shift row contents to the right of the deleted cell to the left
  # if :up is specified, method will shift column contents below the deleted cell upward
  def delete_cell(row_index = 0, column_index=0, shift=nil)
    validate_workbook
    validate_nonnegative(row_index)
    validate_nonnegative(column_index)

    row = sheet_data[row_index]
    old_cell = row && row[column_index]

    case shift
    when nil then
      row.cells[column_index] = nil if row
    when :left then
      row.delete_cell_shift_left(column_index) if row
    when :up then
      (row_index...(sheet_data.size - 1)).each { |index|
        c = sheet_data.rows[index].cells[column_index] = sheet_data.rows[index + 1].cells[column_index]
        c.row -= 1 if c.is_a?(Cell)
      }
    else
      raise 'invalid shift option'
    end

    return old_cell
  end

  private

  NAME = 0
  SIZE = 1
  COLOR = 2
  ITALICS = 3
  BOLD = 4
  UNDERLINE = 5
  STRIKETHROUGH = 6

  def row_font(row)
    (row = sheet_data.rows[row]) && row.get_font
  end

  def get_row_alignment(row, is_horizontal)
    validate_workbook
    validate_nonnegative(row)

    xf_obj = get_row_xf(row)
    return nil if xf_obj.alignment.nil?

    if is_horizontal then return xf_obj.alignment.horizontal
    else                  return xf_obj.alignment.vertical
    end
  end

  def column_font(col)
    validate_workbook
    validate_nonnegative(col)

    @workbook.fonts[@workbook.cell_xfs[get_cols_style_index(col)].font_id]
  end

  #validates Workbook, ensures that this worksheet is in @workbook
  def validate_workbook()
    unless @workbook.nil? || @workbook.worksheets.nil?
      return if @workbook.worksheets.include?(self)
    end

    raise "This worksheet #{self} is not in workbook #{@workbook}"
  end

  def get_cols_style_index(column_index)
    range = cols.locate_range(column_index)
    (range && range.style_index) || 0
  end

  # Helper method to update the row styles array
  # change_type - NAME or SIZE or COLOR etc
  # main method to change font, called from each separate font mutator method
  def change_row_font(row_index, change_type, arg, font)
    validate_workbook
    ensure_cell_exists(row_index)

    xf = workbook.register_new_font(font, get_row_xf(row_index))
    row = sheet_data[row_index]
    row.style_index = workbook.register_new_xf(xf, get_row_style(row_index))
    row.cells.each { |c| c.font_switch(change_type, arg) unless c.nil? }
  end

  # Helper method to update the fonts and cell styles array
  # main method to change font, called from each separate font mutator method
  def change_column_font(column_index, change_type, arg, font, xf)
    validate_workbook
    ensure_cell_exists(0, column_index)

    xf = workbook.register_new_font(font, xf)
    cols.get_range(column_index).style_index = workbook.register_new_xf(xf, get_col_style(column_index))

    sheet_data.rows.each { |row|
      c = row && row[column_index]
      c.font_switch(change_type, arg) unless c.nil?
    }
  end

  # Ensures that cell with +row_index+ and +column_index+ exists in
  #  +sheet_data+ arrays, growing them up if necessary.
  def ensure_cell_exists(row_index, column_index = 0)
    validate_nonnegative(row_index)
    validate_nonnegative(column_index)

    row = sheet_data.rows[row_index] || add_row(row_index)
  end  

  # Helper method to get the style index for a column
  def get_col_style(column_index)
    range = cols.locate_range(column_index)
    (range && range.style_index) || 0
  end

  def get_row_style(row_index)
    row = sheet_data.rows[row_index]
    (row && row.style_index) || 0
  end

  def get_col_xf(column_index)
    @workbook.cell_xfs[get_col_style(column_index)]
  end

  def get_row_xf(row)
    @workbook.cell_xfs[get_row_style(row)]
  end

  def validate_nonnegative(row_or_col)
    raise 'Row and Column arguments must be nonnegative' if row_or_col < 0
  end
  private :validate_nonnegative

end #end class
end
