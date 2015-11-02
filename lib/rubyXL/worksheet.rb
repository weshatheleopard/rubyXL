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

  private

  #validates Workbook, ensures that this worksheet is in @workbook
  def validate_workbook()
    unless @workbook.nil? || @workbook.worksheets.nil?
      return if @workbook.worksheets.include?(self)
    end

    raise "This worksheet #{self} is not in workbook #{@workbook}"
  end

  # Ensures that cell with +row_index+ and +column_index+ exists in
  #  +sheet_data+ arrays, growing them up if necessary.
  def ensure_cell_exists(row_index, column_index = 0)
    validate_nonnegative(row_index)
    validate_nonnegative(column_index)

    row = sheet_data.rows[row_index] || add_row(row_index)
  end  

  def get_col_xf(column_index)
    @workbook.cell_xfs[get_col_style(column_index)]
  end

  def get_row_xf(row)
    validate_nonnegative(row)
    @workbook.cell_xfs[get_row_style(row)]
  end

  def validate_nonnegative(row_or_col)
    raise 'Row and Column arguments must be nonnegative' if row_or_col < 0
  end
  private :validate_nonnegative

end #end class
end
