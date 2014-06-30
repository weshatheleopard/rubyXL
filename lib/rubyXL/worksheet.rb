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

  #changes color of fill in (zer0 indexed) row
  def change_row_fill(row_index = 0, rgb = 'ffffff')
    validate_workbook
    ensure_cell_exists(row_index)
    Color.validate_color(rgb)

    sheet_data.rows[row_index].style_index = @workbook.modify_fill(get_row_style(row_index), rgb)
    sheet_data[row_index].cells.each { |c| c.change_fill(rgb) unless c.nil? }
  end

  def change_row_font_name(row = 0, font_name = 'Verdana')
    ensure_cell_exists(row)
    font = row_font(row).dup
    font.set_name(font_name)
    change_row_font(row, Worksheet::NAME, font_name, font)
  end

  def change_row_font_size(row = 0, font_size=10)
    ensure_cell_exists(row)
    font = row_font(row).dup
    font.set_size(font_size)
    change_row_font(row, Worksheet::SIZE, font_size, font)
  end

  def change_row_font_color(row = 0, font_color = '000000')
    ensure_cell_exists(row)
    Color.validate_color(font_color)
    font = row_font(row).dup
    font.set_rgb_color(font_color)
    change_row_font(row, Worksheet::COLOR, font_color, font)
  end

  def change_row_italics(row = 0, italicized = false)
    ensure_cell_exists(row)
    font = row_font(row).dup
    font.set_italic(italicized)
    change_row_font(row, Worksheet::ITALICS, italicized, font)
  end

  def change_row_bold(row = 0, bolded = false)
    ensure_cell_exists(row)
    font = row_font(row).dup
    font.set_bold(bolded)
    change_row_font(row, Worksheet::BOLD, bolded, font)
  end

  def change_row_underline(row = 0, underlined=false)
    ensure_cell_exists(row)
    font = row_font(row).dup
    font.set_underline(underlined)
    change_row_font(row, Worksheet::UNDERLINE, underlined, font)
  end

  def change_row_strikethrough(row = 0, struckthrough=false)
    ensure_cell_exists(row)
    font = row_font(row).dup
    font.set_strikethrough(struckthrough)
    change_row_font(row, Worksheet::STRIKETHROUGH, struckthrough, font)
  end

  def change_row_height(row = 0, height=10)
    validate_workbook
    ensure_cell_exists(row)

    if height.to_i.to_s == height.to_s
      height = Integer(height)
    elsif height.to_f.to_s == height.to_s
      height = Float(height)
    else
      raise 'You must enter a number for the height'
    end

    sheet_data.rows[row].ht = height
    sheet_data.rows[row].custom_height = true
  end

  def change_row_horizontal_alignment(row = 0,alignment='center')
    validate_workbook
    validate_nonnegative(row)
    change_row_alignment(row,alignment,true)
  end

  def change_row_vertical_alignment(row = 0,alignment='center')
    validate_workbook
    validate_nonnegative(row)
    change_row_alignment(row,alignment,false)
  end

  def change_row_border_top(row = 0, weight = 'thin')
    change_row_border(row, :top, weight)
  end

  def change_row_border_left(row = 0, weight = 'thin')
    change_row_border(row, :left, weight)
  end

  def change_row_border_right(row = 0, weight = 'thin')
    change_row_border(row, :right, weight)
  end

  def change_row_border_bottom(row = 0, weight = 'thin')
    change_row_border(row, :bottom, weight)
  end

  def change_row_border_diagonal(row = 0, weight = 'thin')
    change_row_border(row, :diagonal, weight)
  end

  # Changes font name of column
  def change_column_font_name(column_index = 0, font_name = 'Verdana')
    xf = get_col_xf(column_index)
    font = @workbook.fonts[xf.font_id].dup
    font.set_name(font_name)
    change_column_font(column_index, Worksheet::NAME, font_name, font, xf)
  end

  # Changes font size of column
  def change_column_font_size(column_index, font_size=10)
    xf = get_col_xf(column_index)
    font = @workbook.fonts[xf.font_id].dup
    font.set_size(font_size)
    change_column_font(column_index, Worksheet::SIZE, font_size, font, xf)
  end

  # Changes font color of column
  def change_column_font_color(column_index, font_color='000000')
    Color.validate_color(font_color)

    xf = get_col_xf(column_index)
    font = @workbook.fonts[xf.font_id].dup
    font.set_rgb_color(font_color)
    change_column_font(column_index, Worksheet::COLOR, font_color, font, xf)
  end

  def change_column_italics(column_index, italicized = false)
    xf = get_col_xf(column_index)
    font = @workbook.fonts[xf.font_id].dup
    font.set_italic(italicized)
    change_column_font(column_index, Worksheet::ITALICS, italicized, font, xf)
  end

  def change_column_bold(column_index, bolded = false)
    xf = get_col_xf(column_index)
    font = @workbook.fonts[xf.font_id].dup
    font.set_bold(bolded)
    change_column_font(column_index, Worksheet::BOLD, bolded, font, xf)
  end

  def change_column_underline(column_index, underlined = false)
    xf = get_col_xf(column_index)
    font = @workbook.fonts[xf.font_id].dup
    font.set_underline(underlined)
    change_column_font(column_index, Worksheet::UNDERLINE, underlined, font, xf)
  end

  def change_column_strikethrough(column_index, struckthrough=false)
    xf = get_col_xf(column_index)
    font = @workbook.fonts[xf.font_id].dup
    font.set_strikethrough(struckthrough)
    change_column_font(column_index, Worksheet::STRIKETHROUGH, struckthrough, font, xf)
  end

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

  def change_column_horizontal_alignment(column_index, alignment = 'center')
    change_column_alignment(column_index, alignment,true)
  end

  def change_column_vertical_alignment(column_index, alignment = 'center')
    change_column_alignment(column_index, alignment, false)
  end

  def change_column_border_top(column_index, weight = 'thin')
    change_column_border(column_index, :top, weight)
  end

  def change_column_border_left(column_index, weight = 'thin')
    change_column_border(column_index, :left, weight)
  end

  def change_column_border_right(column_index, weight = 'thin')
    change_column_border(column_index, :right, weight)
  end

  def change_column_border_bottom(column_index, weight = 'thin')
    change_column_border(column_index, :bottom, weight)
  end

  def change_column_border_diagonal(column_index, weight = 'thin')
    change_column_border(column_index, :diagonal, weight)
  end

  # merges cells within a rectangular range
  def merge_cells(row1 = 0, col1 = 0, row2 = 0, col2 = 0)
    validate_workbook

    self.merged_cells ||= RubyXL::MergedCells.new
    merged_cells << RubyXL::MergedCell.new(:ref => RubyXL::Reference.new(row1, row2, col1, col2))
  end

  def add_row(row = 0, params = {})
    new_row = RubyXL::Row.new(params)
    new_row.worksheet = self
    sheet_data.rows[row] = new_row
  end

  def add_cell(row = 0, column = 0, data = '', formula = nil, overwrite = true)
    validate_workbook
    ensure_cell_exists(row, column)

    if overwrite || sheet_data.rows[row].cells[column].nil?
      c = RubyXL::Cell.new
      c.worksheet = self
      c.row = row
      c.column = column
      c.raw_value = data
      c.datatype = RubyXL::DataType::RAW_STRING unless formula || data.is_a?(Numeric)
      c.formula = RubyXL::Formula.new(:expression => formula) if formula
      
      range = cols && cols.find(column)
      c.style_index = sheet_data.rows[row].style_index || (range && range.style_index) || 0

      sheet_data.rows[row].cells[column] = c
    end

    sheet_data.rows[row].cells[column]
  end

  def delete_row(row_index=0)
    validate_workbook
    validate_nonnegative(row_index)
    return nil unless row_exists(row_index)

    deleted = sheet_data.rows.delete_at(row_index)
    row_num = row_index+1

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

    return nil unless column_exists(column_index)

    sheet_data.rows.each { |row| row.cells.delete_at(column_index) }

    # Change column numbers for cells to the right of the deleted column
    sheet_data.rows.each_with_index { |row, row_index|
      row.cells.each_with_index { |c, column_index|
        c.column = column_index if c.is_a?(Cell)
      }
    }

    cols.column_ranges.each { |range| range.delete_column(column_index) }
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
  def delete_cell(row = 0, col=0, shift=nil)
    validate_workbook
    validate_nonnegative(row)
    validate_nonnegative(col)

    return nil unless row_exists(row) && column_exists(col)

    cell = sheet_data[row][col]

    case shift
    when nil then
      sheet_data.rows[row].cells[col] = nil
    when :left then
      sheet_data.rows[row].delete_cell_shift_left(col)
    when :up then
      (row...(sheet_data.size - 1)).each { |index|
        c = sheet_data.rows[index].cells[col] = sheet_data.rows[index + 1].cells[col]
        c.row -= 1 if c.is_a?(Cell)
      }
    else
      raise 'invalid shift option'
    end

    return cell
  end

  def get_row_fill(row = 0)
    (row = sheet_data.rows[row]) && row.get_fill_color
  end

  def get_row_font_name(row = 0)
    (font = row_font(row)) && font.get_name
  end

  def get_row_font_size(row = 0)
    (font = row_font(row)) && font.get_size
  end

  def get_row_font_color(row = 0)
    font = row_font(row)
    color = font && font.color
    color && (color.rgb || '000000')
  end

  def is_row_italicized(row = 0)
    (font = row_font(row)) && font.is_italic
  end

  def is_row_bolded(row = 0)
    (font = row_font(row)) && font.is_bold
  end

  def is_row_underlined(row = 0)
    (font = row_font(row)) && font.is_underlined
  end

  def is_row_struckthrough(row = 0)
    (font = row_font(row)) && font.is_strikethrough
  end

  def get_row_height(row = 0)
    validate_workbook
    validate_nonnegative(row)
    return nil unless row_exists(row)
    row = sheet_data.rows[row]
    row && row.ht || 13
  end

  def get_row_horizontal_alignment(row = 0)
    return get_row_alignment(row, true)
  end

  def get_row_vertical_alignment(row = 0)
    return get_row_alignment(row, false)
  end

  def get_row_border_top(row = 0)
    return get_row_border(row, :top)
  end

  def get_row_border_left(row = 0)
    return get_row_border(row, :left)
  end                         

  def get_row_border_right(row = 0)
    return get_row_border(row, :right)
  end

  def get_row_border_bottom(row = 0)
    return get_row_border(row, :bottom)
  end

  def get_row_border_diagonal(row = 0)
    return get_row_border(row, :diagonal)
  end

  def get_column_font_name(col = 0)
    font = column_font(col)
    font && font.get_name
  end

  def get_column_font_size(col = 0)
    font = column_font(col)
    font && font.get_size
  end

  def get_column_font_color(col = 0)
    font = column_font(col)
    font && (font.get_rgb_color || '000000')
  end

  def is_column_italicized(col = 0)
    font = column_font(col)
    font && font.is_italic
  end

  def is_column_bolded(col = 0)
    font = column_font(col)
    font && font.is_bold
  end

  def is_column_underlined(col = 0)
    font = column_font(col)
    font && font.is_underlined
  end

  def is_column_struckthrough(col = 0)
    font = column_font(col)
    font && font.is_strikethrough
  end

  # Get raw column width value as stored in the file
  def get_column_width_raw(column_index = 0)
    validate_workbook
    validate_nonnegative(column_index)
    return nil unless column_exists(column_index)

    range = cols.find(column_index)
    range && range.width
  end

  # Get column width measured in number of digits, as per
  # http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column%28v=office.14%29.aspx
  def get_column_width(column_index = 0)
    width = get_column_width_raw(column_index)
    return RubyXL::ColumnRange::DEFAULT_WIDTH if width.nil?
    (width - (5.0 / RubyXL::Font::MAX_DIGIT_WIDTH)).round
  end

  def get_column_fill(col=0)
    validate_workbook
    validate_nonnegative(col)
    return nil unless column_exists(col)

    @workbook.get_fill_color(get_col_xf(col))
  end

  def get_column_horizontal_alignment(col=0)
    get_column_alignment(col, :horizontal)
  end

  def get_column_vertical_alignment(col=0)
    get_column_alignment(col, :vertical)
  end

  def get_column_border_top(col=0)
    get_column_border(col, :top)
  end

  def get_column_border_left(col=0)
    get_column_border(col, :left)
  end

  def get_column_border_right(col=0)
    get_column_border(col, :right)
  end

  def get_column_border_bottom(col=0)
    get_column_border(col, :bottom)
  end

  def get_column_border_diagonal(col=0)
    get_column_border(col, :diagonal)
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

    return nil unless row_exists(row)

    xf_obj = get_row_xf(row)
    return nil if xf_obj.alignment.nil?

    if is_horizontal then return xf_obj.alignment.horizontal
    else                  return xf_obj.alignment.vertical
    end
  end

  def get_row_border(row, border_direction)
    validate_workbook
    validate_nonnegative(row)

    return nil unless row_exists(row)

    border = @workbook.borders[get_row_xf(row).border_id]
    border && border.get_edge_style(border_direction)
  end

  def column_font(col)
    validate_workbook
    validate_nonnegative(col)
    return nil unless column_exists(col)

    @workbook.fonts[@workbook.cell_xfs[get_cols_style_index(col)].font_id]
  end

  def get_column_alignment(col, type)
    validate_workbook
    validate_nonnegative(col)
    return nil unless column_exists(col)

    xf = @workbook.cell_xfs[get_cols_style_index(col)]
    xf.alignment && xf.alignment.send(type)
  end

  def get_column_border(col, border_direction)
    validate_workbook
    validate_nonnegative(col)
    return nil unless column_exists(col)

    xf = @workbook.cell_xfs[get_cols_style_index(col)]
    border = @workbook.borders[xf.border_id]
    border && border.get_edge_style(border_direction)
  end

  #validates Workbook, ensures that this worksheet is in @workbook
  def validate_workbook()
    unless @workbook.nil? || @workbook.worksheets.nil?
      return if @workbook.worksheets.include?(self)
    end

    raise "This worksheet #{self} is not in workbook #{@workbook}"
  end

  def get_cols_style_index(column_index)
    range = cols.find(column_index)
    (range && range.style_index) || 0
  end

  # Helper method to update the row styles array
  # change_type - NAME or SIZE or COLOR etc
  # main method to change font, called from each separate font mutator method
  def change_row_font(row, change_type, arg, font)
    validate_workbook
    ensure_cell_exists(row)

    xf = workbook.register_new_font(font, get_row_xf(row))
    sheet_data.rows[row].style_index = workbook.register_new_xf(xf, get_row_style(row))

    sheet_data[row].cells.each { |c| font_switch(c, change_type, arg) unless c.nil? }
  end

  # Helper method to update the fonts and cell styles array
  # main method to change font, called from each separate font mutator method
  def change_column_font(column_index, change_type, arg, font, xf)
    validate_workbook
    ensure_cell_exists(0, column_index)

    xf = workbook.register_new_font(font, xf)

    cols.get_range(column_index).style_index = workbook.register_new_xf(xf, get_col_style(column_index))

    sheet_data.rows.each { |row|
      c = row[column_index]
      font_switch(c, change_type, arg) unless c.nil?
    }
  end

  #performs correct modification based on what type of change_type is specified
  def font_switch(c,change_type,arg)
    case change_type
      when Worksheet::NAME
        unless arg.is_a?String
          raise 'Not a String'
        end
        c.change_font_name(arg)
      when Worksheet::SIZE
        unless arg.is_a?(Integer) || arg.is_a?(Float)
          raise 'Not a Number'
        end
          c.change_font_size(arg)
      when Worksheet::COLOR
        Color.validate_color(arg)
        c.change_font_color(arg)
      when Worksheet::ITALICS
        unless arg == !!arg
          raise 'Not a boolean'
        end
        c.change_font_italics(arg)
      when Worksheet::BOLD
        unless arg == !!arg
          raise 'Not a boolean'
        end
        c.change_font_bold(arg)
      when Worksheet::UNDERLINE
        unless arg == !!arg
          raise 'Not a boolean'
        end
        c.change_font_underline(arg)
      when Worksheet::STRIKETHROUGH
        unless arg == !!arg
          raise 'Not a boolean'
        end
        c.change_font_strikethrough(arg)
      else
        raise 'Invalid change_type'
    end
  end

  # Ensures that cell with +row_index+ and +column_index+ exists in
  #  +sheet_data+ arrays, growing them up if necessary.
  def ensure_cell_exists(row_index, column_index = 0)
    validate_nonnegative(row_index)
    validate_nonnegative(column_index)

    existing_row_count = sheet_data.rows.size

    # Expand cell arrays in existing rows, if necessary.
    # Writing anything to a cell in the array automatically creates all the members
    # with lower indices, filling them with +nil+s. But, we can't just write +nil+
    # to +column_index+ because it may be less than +size+! So we read from that index
    # (if it didn't exist, we will get nil) and write right back.
    sheet_data.rows.each { |r| r.cells[column_index] = r.cells[column_index] unless r.nil? }

    first_row = sheet_data.rows.first
    col_size = [ first_row && first_row.cells.size || 0, column_index ].max

    # Now create new rows with the required number of cells.
    # Doing +.downto()+ here so the reallocation of row array has to only happen once,
    # when it is extended to max size; after that, we will be writing into existing
    # (but empty) members. Additional checks are not necessary, because if +row_index+
    # is less than +size+, then +.downto()+ will not execute, and if it equals +size+,
    # then the block will be invoked exactly once, which takes care of the case when
    # +row_index+ is greater than the current max index by exactly 1.
    row_index.downto(existing_row_count) { |r| 
      add_row(r, :cells => Array.new(col_size))
    } 
  end  

  # Helper method to get the style index for a column
  def get_col_style(column_index)
    range = cols.find(column_index)
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

  def change_row_alignment(row,alignment,is_horizontal)
    validate_workbook
    validate_nonnegative(row)
    ensure_cell_exists(row)

    sheet_data.rows[row].style_index = @workbook.modify_alignment(get_row_style(row), is_horizontal, alignment)

    sheet_data[row].cells.each { |c|
      next if c.nil?
      if is_horizontal then c.change_horizontal_alignment(alignment)
      else                  c.change_vertical_alignment(alignment)
      end
    }
  end

  def change_column_alignment(column_index, alignment, is_horizontal)
    validate_workbook
    ensure_cell_exists(0, column_index)

    cols.get_range(column_index).style_index = @workbook.modify_alignment(get_col_style(column_index), is_horizontal, alignment)
    # Excel gets confused if width is not explicitly set for a column that had alignment changes
    change_column_width(column_index) if get_column_width_raw(column_index).nil?

    sheet_data.rows.each { |row|
      c = row[column_index]
      next if c.nil?
      if is_horizontal
        c.change_horizontal_alignment(alignment)
      else
        c.change_vertical_alignment(alignment)
      end
    }
  end

  def change_row_border(row, direction, weight)
    validate_workbook
    ensure_cell_exists(row)

    sheet_data.rows[row].style_index = @workbook.modify_border(get_row_style(row), direction, weight)

    sheet_data[row].cells.each { |c|
      c.change_border(direction, weight) unless c.nil?
    }
  end

  def change_column_border(column_index, direction, weight)
    validate_workbook
    ensure_cell_exists(0, column_index)

    cols.get_range(column_index).style_index = @workbook.modify_border(get_col_style(column_index), direction, weight)

    sheet_data.rows.each { |row|
      c = row.cells[column_index]
      c.change_border(direction, weight) unless c.nil?
    }
  end

  def validate_nonnegative(row_or_col)
    raise 'Row and Column arguments must be nonnegative' if row_or_col < 0
  end
  private :validate_nonnegative

  def column_exists(col)
    sheet_data.rows[0].cells.size > col
  end

  def row_exists(row)
    sheet_data.rows.size > row
  end


end #end class
end
