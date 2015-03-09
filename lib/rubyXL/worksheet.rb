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
    row_num = find_first_row_with_content(headers)
    return nil if row_num.nil?

    table_hash = {}
    table_hash[:table] = []

    header_row = sheet_data[row_num]
    header_row.cells.each_with_index { |header_cell, index|
      break if index>0 && !opts[:last_header].nil? && !header_row[index-1].nil? && !header_row[index-1].value.nil? && header_row[index-1].value.to_s==opts[:last_header]
      next if header_cell.nil? || header_cell.value.nil?
      header = header_cell.value.to_s
      table_hash[:sorted_headers]||=[]
      table_hash[:sorted_headers] << header
      table_hash[header] = []

      original_row = row_num + 1
      current_row = original_row

      row = sheet_data.rows[current_row]
      cell = row && row.cells[index]

      # makes array of hashes in table_hash[:table]
      # as well as hash of arrays in table_hash[header]
      table_index = current_row - original_row
      cell_test = (!cell.nil? && !cell.value.nil?)

      while cell_test || (table_hash[:table][table_index] && !table_hash[:table][table_index].empty?)
        table_hash[header] << cell.value if cell_test
        table_index = current_row - original_row

        if cell_test then
          table_hash[:table][table_index] ||= {}
          table_hash[:table][table_index][header] = cell.value
        end

        current_row += 1
        if sheet_data.rows[current_row].nil? then
          cell = nil
        else
          cell = sheet_data.rows[current_row].cells[index]
        end
        cell_test = (!cell.nil? && !cell.value.nil?)
      end
    }

    return table_hash
  end

  # finds first row which contains at least all strings in cells_content
  def find_first_row_with_content(cells_content)
    validate_workbook

    sheet_data.rows.each_with_index { |row, index|
      next if row.nil?
      cells_content = cells_content.map { |header| header.to_s.strip.downcase }
      original_cells_content = row.cells.map { |cell| (cell && cell.value).to_s.strip.downcase }

      if (cells_content & original_cells_content).size == cells_content.size
        return index
      end
    }
    return nil
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
                      else nc = RubyXL::Cell.new(:style_index => c.style_index)
                           nc.worksheet = self
                           nc
                      end
                    }
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
