module RubyXL
class Worksheet
  include Enumerable

  attr_accessor :sheet_name, :sheet_id, :sheet_data, :column_ranges, :merged_cells, :pane,
                :validations, :sheet_views, :legacy_drawings, :extLst, :workbook,
                :row_styles, :drawings, :sheet_data2

  SHEET_NAME_TEMPLATE = 'Sheet%d'

  def initialize(workbook, sheet_name = nil, sheet_data= [[nil]], cols=[], merged_cells=[])
    @workbook = workbook

    @sheet_name = sheet_name || get_default_name
    @sheet_id = nil
    @sheet_data = sheet_data
    @sheet_data2 = RubyXL::SheetData.new
    @column_ranges = cols
    @merged_cells = merged_cells || []
    @row_styles = []
    @sheet_views = [ RubyXL::SheetView.new ]
    @extLst = nil
    @legacy_drawings = []
    @drawings = []
    @validations = []
  end

  def get_default_name
    n = 0

    begin
      name = SHEET_NAME_TEMPLATE % (n += 1)
    end until @workbook[name].nil?

    name
  end
  private :get_default_name

  # allows for easier access to sheet_data
  def [](row = 0)
    @sheet_data[row]
  end

  def each
    @sheet_data.each { |row| yield(row) }
  end

  #returns 2d array of just the cell values (without style or formula information)
  def extract_data(args = {})
    @sheet_data.map {|row| row.map {|c| if c.is_a?(Cell) then c.value(args) else nil end}}
  end

  def get_table(headers = [], opts = {})
    validate_workbook

    headers = [headers] unless headers.is_a?(Array)
    row_num = find_first_row_with_content(headers)
  	
    return nil if row_num.nil?

    table_hash = {}
    table_hash[:table] = []

    header_row = @sheet_data[row_num]
    header_row.each_with_index { |header_cell, index|
      break if index>0 && !opts[:last_header].nil? && !header_row[index-1].nil? && !header_row[index-1].value.nil? && header_row[index-1].value.to_s==opts[:last_header]
      next if header_cell.nil? || header_cell.value.nil?
      header = header_cell.value.to_s
      table_hash[:sorted_headers]||=[]
      table_hash[:sorted_headers] << header
      table_hash[header] = []

      original_row = row_num + 1
      current_row = original_row

      cell = @sheet_data[current_row][index]

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
        if @sheet_data[current_row].nil? then
          cell = nil
        else
          cell = @sheet_data[current_row][index]
        end
        cell_test = (!cell.nil? && !cell.value.nil?)
      end
    }

    return table_hash
  end

  #changes color of fill in (zer0 indexed) row
  def change_row_fill(row = 0, rgb = 'ffffff')
    validate_workbook
    ensure_cell_exists(row)
    Color.validate_color(rgb)

    if @row_styles[(Integer(row)+1)].nil?
      @row_styles[(Integer(row)+1)] = {}
      @row_styles[(Integer(row)+1)][:style] = 0
    end

    @row_styles[(Integer(row)+1)][:style] = @workbook.modify_fill(Integer(@row_styles[(Integer(row)+1)][:style]),rgb)

    @sheet_data[Integer(row)].each do |c|
      unless c.nil?
        c.change_fill(rgb)
      end
    end
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

  def change_row_font_color(row = 0, font_color='000000')
    ensure_cell_exists(row)
    Color.validate_color(font_color)
    font = row_font(row).dup
    font.set_rgb_color(font_color)
    change_row_font(row, Worksheet::COLOR, font_color, font)
  end

  def change_row_italics(row = 0, italicized=false)
    ensure_cell_exists(row)
    font = row_font(row).dup
    font.set_italic(italicized)
    change_row_font(row, Worksheet::ITALICS, italicized, font)
  end

  def change_row_bold(row = 0, bolded=false)
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

    if @row_styles[(row+1)].nil?
      @row_styles[(row+1)] = {}
      @row_styles[(row+1)][:style] = 0
    end
    @row_styles[(row+1)][:height] = height
    @row_styles[(row+1)][:customHeight] = '1'
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
  def change_column_font_name(col = 0, font_name = 'Verdana')
    xf = get_col_xf(col)
    font = @workbook.fonts[xf.font_id].dup
    font.set_name(font_name)
    change_column_font(col, Worksheet::NAME, font_name, font, xf)
  end

  # Changes font size of column
  def change_column_font_size(col=0, font_size=10)
    xf = get_col_xf(col)
    font = @workbook.fonts[xf.font_id].dup
    font.set_size(font_size)
    change_column_font(col, Worksheet::SIZE, font_size, font, xf)
  end

  # Changes font color of column
  def change_column_font_color(col=0, font_color='000000')
    Color.validate_color(font_color)

    xf = get_col_xf(col)
    font = @workbook.fonts[xf.font_id].dup
    font.set_rgb_color(font_color)
    change_column_font(col, Worksheet::COLOR, font_color, font, xf)
  end

  def change_column_italics(col = 0, italicized = false)
    xf = get_col_xf(col)
    font = @workbook.fonts[xf.font_id].dup
    font.set_italic(italicized)
    change_column_font(col, Worksheet::ITALICS, italicized, font, xf)
  end

  def change_column_bold(col = 0, bolded = false)
    xf = get_col_xf(col)
    font = @workbook.fonts[xf.font_id].dup
    font.set_bold(bolded)
    change_column_font(col, Worksheet::BOLD, bolded, font, xf)
  end

  def change_column_underline(col = 0, underlined = false)
    xf = get_col_xf(col)
    font = @workbook.fonts[xf.font_id].dup
    font.set_underline(underlined)
    change_column_font(col, Worksheet::UNDERLINE, underlined, font, xf)
  end

  def change_column_strikethrough(col=0, struckthrough=false)
    xf = get_col_xf(col)
    font = @workbook.fonts[xf.font_id].dup
    font.set_strikethrough(struckthrough)
    change_column_font(col, Worksheet::STRIKETHROUGH, struckthrough, font, xf)
  end

  def change_column_width(col = 0, width = 13)
    validate_workbook
    ensure_cell_exists(0, col)

    RubyXL::ColumnRange.update(col, @column_ranges, { :width => width, :custom_width => 1 })
  end

  def get_column_style_index(col)
    range = RubyXL::ColumnRange.find(col, @column_ranges)
    (range && range.style_index) || 0
  end

  def change_column_fill(col=0, color_index='ffffff')
    validate_workbook
    Color.validate_color(color_index)
    ensure_cell_exists(0, col)

    new_style_index = @workbook.modify_fill(get_column_style_index(col), color_index)

    RubyXL::ColumnRange.update(col, @column_ranges, { :style => new_style_index })

    @sheet_data.each { |row|
      c = row[col]
      c.change_fill(color_index) if c
    }
  end

  def change_column_horizontal_alignment(col=0,alignment='center')
    change_column_alignment(col,alignment,true)
  end

  def change_column_vertical_alignment(col=0,alignment='center')
    change_column_alignment(col, alignment, false)
  end

  def change_column_border_top(col=0,weight = 'thin')
    change_column_border(col, :top, weight)
  end

  def change_column_border_left(col=0,weight = 'thin')
    change_column_border(col, :left, weight)
  end

  def change_column_border_right(col=0,weight = 'thin')
    change_column_border(col, :right, weight)
  end

  def change_column_border_bottom(col=0,weight = 'thin')
    change_column_border(col, :bottom, weight)
  end

  def change_column_border_diagonal(col=0,weight = 'thin')
    change_column_border(col, :diagonal, weight)
  end

  # merges cells within a rectangular range
  def merge_cells(row1 = 0, col1 = 0, row2 = 0, col2 = 0)
    validate_workbook
    @merged_cells << RubyXL::Reference.new(row1, row2, col1, col2)
  end

  def add_cell(row = 0, column = 0, data='', formula=nil, overwrite=true)
    validate_workbook
    ensure_cell_exists(row, column)

    if overwrite || @sheet_data[row][column].nil?
      c = Cell.new
      c.worksheet = self
      c.row = row
      c.column = column
      c.raw_value = data
      c.datatype = (formula || data.is_a?(Numeric)) ? '' : RubyXL::Cell::RAW_STRING
      c.formula = formula
      c.style_index = 0

      @sheet_data[row][column] = c

      col = RubyXL::ColumnRange.find(column, @column_ranges)

      if @row_styles[(row+1)] != nil
        @sheet_data[row][column].style_index = @row_styles[(row+1)][:style]
      elsif col != nil
        @sheet_data[row][column].style_index = col.style_index
      end
    end

    add_cell_style(row, column)

    return @sheet_data[row][column]
  end

  def add_cell_obj(cell, overwrite=true)
    validate_workbook

    return nil if cell.nil?

    row = cell.row
    column = cell.column

    ensure_cell_exists(row, column)

    if overwrite || @sheet_data[row][column].nil?
      @sheet_data[row][column] = cell
    end

    add_cell_style(row,column)

    return @sheet_data[row][column]
  end

  def delete_row(row_index=0)
    validate_workbook
    validate_nonnegative(row_index)

    if row_index >= @sheet_data.size
      return nil
    end

    deleted = @sheet_data.delete_at(row_index)
    row_num = row_index+1

    @row_styles.delete_at(row_index)

    # Change cell row numbers
    row_index.upto(@sheet_data.size - 1) { |index|
      @sheet_data[index].each{ |c| c.row -= 1 unless c.nil? }
    }

    return deleted
  end

  #inserts row at row_index, pushes down, copies style from below (row previously at that index)
  #USE OF THIS METHOD will break formulas which reference cells which are being "pushed down"
  def insert_row(row_index=0)
    validate_workbook
    ensure_cell_exists(row_index)

    @sheet_data.insert(row_index,Array.new(@sheet_data[row_index].size))

    row_num = row_index+1

    #copy cell styles from row above, (or below if first row)
    @sheet_data[row_index].each_index do |i|
      if row_index > 0
        old_cell = @sheet_data[row_index-1][i]
      else
        old_cell = @sheet_data[row_index+1][i]
      end

      unless old_cell.nil?
        #only add cell if style exists, not copying content

        if @row_styles[(row_num+1)].nil?
          @row_styles[(row_num+1)] = {:style=>0}
        end
        if old_cell.style_index != 0 && old_cell.style_index != @row_styles[(row_num+1)][:style]
          c = Cell.new
          c.worksheet = self
          c.row = row_index
          c.column = i
          c.datatype = RubyXL::Cell::SHARED_STRING
          c.style_index = old_cell.style_index
          @sheet_data[row_index][i] = c
        end
      end
    end

    #copy row styles from row above, (or below if first row)
    (@row_styles.size+1).downto(row_num+1) do |i|
      @row_styles[i] = @row_styles[(i-1)]
    end

    if row_index > 0
      @row_styles[row_num] = @row_styles[(row_num-1)]
    else
      @row_styles[row_num] = nil #@row_styles[(row_num+1).to_s]
    end

    #update row value for all rows below
    (row_index+1).upto(@sheet_data.size-1) do |i|
      row = @sheet_data[i]
      row.each do |c|
        unless c.nil?
          c.row += 1
        end
      end
    end

    return @sheet_data[row_index]
  end

  def delete_column(col_index=0)
    validate_workbook
    validate_nonnegative(col_index)

    if col_index >= @sheet_data[0].size
      return nil
    end

    #delete column
    @sheet_data.map {|r| r.delete_at(col_index)}

    #change column numbers for cells to right of deleted column
    @sheet_data.each_with_index do |row,row_index|
      (col_index...(row.size)).each do |index|
        if @sheet_data[row_index][index].is_a?(Cell)
          @sheet_data[row_index][index].column -= 1
        end
      end
    end

    @column_ranges.each { |range| range.delete_column(col_index) }
  end

  # inserts column at col_index, pushes everything right, takes styles from column to left
  # USE OF THIS METHOD will break formulas which reference cells which are being "pushed down"
  def insert_column(col_index = 0)
    validate_workbook
    ensure_cell_exists(0, col_index)

    old_range = col_index > 0 ? RubyXL::ColumnRange.find(col_index, @column_ranges) : RubyXL::ColumnRange.new

    #go through each cell in column
    @sheet_data.each_with_index do |row, row_index|
      old_cell = row[col_index]
      c = nil

      if old_cell && old_cell.style_index != 0 &&
           old_range && old_range.style != old_cell.style_index then

        c = Cell.new
        c.worksheet = self
        c.row = row_index
        c.column = col_index
        c.datatype = RubyXL::Cell::SHARED_STRING
        c.style_index = old_cell.style_index
      end

      row.insert(col_index, c)
    end

    ColumnRange.insert_column(col_index, @column_ranges)

    #update column numbers
    @sheet_data.each { |row|
      (col_index + 1).upto(row.size) { |col|
        row[col].column = col unless row[col].nil?
      }
    }

  end

  def insert_cell(row = 0, col = 0, data = nil, formula = nil, shift = nil)
    validate_workbook
    ensure_cell_exists(row, col)

    case shift
    when nil then # No shifting at all
    when :right then
      @sheet_data[row].insert(col,nil)
      (row...(@sheet_data[row].size)).each { |index|
        if @sheet_data[row][index].is_a?(Cell)
          @sheet_data[row][index].column += 1
        end
      }
    when :down then
      @sheet_data << Array.new(@sheet_data[row].size)
      (@sheet_data.size-1).downto(row+1) { |index|
        @sheet_data[index][col] = @sheet_data[index-1][col]
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

    return nil if @sheet_data.size <= row || @sheet_data[row].size <= col

    cell = @sheet_data[row][col]

    case shift
    when nil then
      @sheet_data[row][col] = nil
    when :left then
      @sheet_data[row].delete_at(col)
      @sheet_data[row] << nil
      (col...(@sheet_data[row].size)).each { |index|
        if @sheet_data[row][index].is_a?(Cell)
          @sheet_data[row][index].column -= 1
        end
      }
    when :up then
      (row...(@sheet_data.size-1)).each { |index|
        @sheet_data[index][col] = @sheet_data[index+1][col]
        if @sheet_data[index][col].is_a?(Cell)
          @sheet_data[index][col].row -= 1
        end
      }

      @sheet_data.last[col].row -= 1 if @sheet_data.last[col].is_a?(Cell)
    else
      raise 'invalid shift option'
    end

    return cell
  end

  def get_row_fill(row = 0)
    validate_workbook
    validate_nonnegative(row)

    if @sheet_data.size <= row
      return nil
    end

    if @row_styles[(row+1)].nil?
      return "ffffff" #default, white
    end

    xf = get_row_xf(row)

    return @workbook.get_fill_color(xf)
  end

  def get_row_font_name(row = 0)
    font = row_font(row)
    font && font.get_name
  end

  def get_row_font_size(row = 0)
    font = row_font(row)
    font && font.get_size
  end

  def get_row_font_color(row = 0)
    font = row_font(row)
    color = font && font.color
    color && (color.rgb || '000000')
  end

  def is_row_italicized(row = 0)
    font = row_font(row)
    font && font.is_italic
  end

  def is_row_bolded(row = 0)
    font = row_font(row)
    font && font.is_bold
  end

  def is_row_underlined(row = 0)
    font = row_font(row)
    font && font.is_underlined
  end

  def is_row_struckthrough(row = 0)
    font = row_font(row)
    font && font.is_strikethrough
  end

  def get_row_height(row = 0)
    validate_workbook
    validate_nonnegative(row)

    if @sheet_data.size <= row
      return nil
    end

    if @row_styles[(row+1)].nil?
      return 13
    else
      @row_styles[(row+1)][:height]
    end
  end

  def get_row_horizontal_alignment(row = 0)
    return get_row_alignment(row,true)
  end

  def get_row_vertical_alignment(row = 0)
    return get_row_alignment(row,false)
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

  def get_column_width(col=0)
    validate_workbook
    validate_nonnegative(col)

    if @sheet_data[0].size <= col
      return nil
    end

    range = RubyXL::ColumnRange.find(col, @column_ranges)

    (range && range.width) || 10
  end

  def get_column_fill(col=0)
    validate_workbook
    validate_nonnegative(col)
    return nil if @sheet_data[0].size <= col
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

  Worksheet::NAME = 0
  Worksheet::SIZE = 1
  Worksheet::COLOR = 2
  Worksheet::ITALICS = 3
  Worksheet::BOLD = 4
  Worksheet::UNDERLINE = 5
  Worksheet::STRIKETHROUGH = 6

  #row_styles is assumed to not be nil at specified row
  def get_row_xf(row)
    @row_styles[(row+1)] ||= { :style => 0 }
    @workbook.cell_xfs[@row_styles[(row+1)][:style]]
  end

  def row_font(row)
    validate_workbook
    validate_nonnegative(row)
    xf = get_row_xf(row)
    return nil if @sheet_data.size <= row
    @workbook.fonts[xf.font_id]
  end

  def get_row_alignment(row, is_horizontal)
    validate_workbook
    validate_nonnegative(row)

    if @sheet_data.size <= row || @row_styles[(row+1)].nil?
      return nil
    end

    xf_obj = @workbook.cell_xfs[@row_styles[(row+1)][:style]]

    return nil if xf_obj.alignment.nil?

    if is_horizontal then
      return xf_obj.alignment.horizontal
    else
      return xf_obj.alignment.vertical
    end
  end

  def get_row_border(row, border_direction)
    validate_workbook
    validate_nonnegative(row)

    return nil if @sheet_data.size <= row || @row_styles[(row+1)].nil?

    border = @workbook.borders[get_row_xf(row).border_id]
    border && border.get_edge_style(border_direction)
  end

  def column_font(col)
    validate_workbook
    validate_nonnegative(col)

    return nil if @sheet_data[0].size <= col
    style_index = get_cols_style_index(col)
    @workbook.fonts[@workbook.cell_xfs[style_index].font_id]
  end

  def get_column_alignment(col, type)
    validate_workbook
    validate_nonnegative(col)

    return nil if @sheet_data[0].size <= col
    xf = @workbook.cell_xfs[get_cols_style_index(col)]
    xf.alignment && xf.alignment.send(type)
  end

  def get_column_border(col, border_direction)
    validate_workbook
    validate_nonnegative(col)

    return nil if @sheet_data[0].size <= col

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

  def get_cols_style_index(col)
    range = RubyXL::ColumnRange.find(col, @column_ranges)
    (range && range.style) || 0
  end

  # Helper method to update the row styles array
  # change_type - NAME or SIZE or COLOR etc
  # main method to change font, called from each separate font mutator method
  def change_row_font(row, change_type, arg, font)
    validate_workbook
    ensure_cell_exists(row)

    xf = workbook.register_new_font(font, get_row_xf(row))
    @row_styles[(row+1)][:style] = workbook.register_new_xf(xf, @row_styles[(row+1)][:style])

    @sheet_data[row] ||= []
    @sheet_data[Integer(row)].each { |c|
      font_switch(c, change_type, arg) unless c.nil?
    }
  end

  # Helper method to update the fonts and cell styles array
  # main method to change font, called from each separate font mutator method
  def change_column_font(col, change_type, arg, font, xf)
    validate_workbook
    ensure_cell_exists(0, col)

    xf = workbook.register_new_font(font, xf)
    new_style_index = workbook.register_new_xf(xf, get_col_style(col))
    RubyXL::ColumnRange.update(col, @column_ranges, { :style => new_style_index })

    @sheet_data.each { |row|
      c = row[col]
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

  # Ensures that cell with +row_index+ and +col_index+ exists in
  #  +sheet_data+ arrays, growing them up if necessary.
  def ensure_cell_exists(row_index, col_index = 0)
    validate_nonnegative(row_index)
    validate_nonnegative(col_index)

    # Writing anything to a cell in the array automatically creates all the members
    # with lower indices, filling them with +nil+s. But, we can't just write +nil+
    # to +col_index+ because it may be less than +size+! So we read from that index
    # (if it didn't exist, we will get nil) and write right back.
    @sheet_data.each { |r| r[col_index] = r[col_index] }

    col_size = @sheet_data[0].size

    # Doing +.downto()+ here so the reallocation of row array has to only happen once,
    # when it is extended to max size; after that, we will be writing into existing
    # (but empty) members. Additional checks are not necessary, because if +row_index+
    # is less than +size+, then +.downto()+ will not execute, and if it equals +size+,
    # then the block will be invoked exactly once, which takes care of the case when
    # +row_index+ is greater than the current max index by exactly 1.
    row_index.downto(@sheet_data.size) { |r| @sheet_data[r] = Array.new(col_size) } 
  end  

  # Helper method to get the style index for a row
  def get_row_style(row)
    if @row_styles[(row+1)].nil?
      @row_styles[(row+1)] = {}
      @row_styles[(row+1)][:style] = 0
      @workbook.fonts[0].count += 1
    end
    return @row_styles[(row+1)][:style]
  end

  # Helper method to get the style index for a column
  def get_col_style(col)
    range = RubyXL::ColumnRange.find(col, @column_ranges)
    (range && range.style) || 0
  end

  def get_col_xf(col)
    @workbook.cell_xfs[get_col_style(col)]
  end

  def change_row_alignment(row,alignment,is_horizontal)
    validate_workbook
    validate_nonnegative(row)
    ensure_cell_exists(row)

    if @row_styles[(row+1)].nil?
      @row_styles[(row+1)] = {}
      @row_styles[(row+1)][:style] = 0
    end

    @row_styles[(row+1)][:style] =
      @workbook.modify_alignment(@row_styles[(row+1)][:style], is_horizontal, alignment)

    @sheet_data[row].each do |c|
      unless c.nil?
        if is_horizontal
          c.change_horizontal_alignment(alignment)
        else
          c.change_vertical_alignment(alignment)
        end
      end
    end
  end

  def change_column_alignment(col,alignment,is_horizontal)
    validate_workbook
    ensure_cell_exists(0, col)

    new_style_index = @workbook.modify_alignment(get_column_style_index(col), is_horizontal, alignment)
    RubyXL::ColumnRange.update(col, @column_ranges, { :style => new_style_index })

    @sheet_data.each { |row|
      c = row[col]
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

    xf = get_row_xf(row)
    border = @workbook.borders[xf.border_id].dup
    border.set_edge_style(direction, weight)

    xf = workbook.register_new_border(border, xf)
    @row_styles[(row+1)][:style] = workbook.register_new_xf(xf, @row_styles[(row+1)][:style])

    @sheet_data[row].each { |c|
      next if c.nil?
      case direction
      when :top      then c.change_border_top(weight)
      when :left     then c.change_border_left(weight)
      when :right    then c.change_border_right(weight)
      when :bottom   then c.change_border_bottom(weight)
      when :diagonal then c.change_border_diagonal(weight)
      else raise 'invalid direction'
      end
    }
  end

  def change_column_border(col, direction, weight)
    col = Integer(col)
    validate_workbook
    ensure_cell_exists(0, col)
     
    xf = get_col_xf(col)
    border = @workbook.borders[xf.border_id].dup
    border.set_edge_style(direction, weight)

    xf = workbook.register_new_border(border, xf)
    new_style_index = workbook.register_new_xf(xf, get_col_style(col))
    RubyXL::ColumnRange.update(col, @column_ranges, { :style => new_style_index })

    @sheet_data.each { |row|
      c = row[col]
      next if c.nil?
      case direction
      when :top      then c.change_border_top(weight)
      when :left     then c.change_border_left(weight)
      when :right    then c.change_border_right(weight)
      when :bottom   then c.change_border_bottom(weight)
      when :diagonal then c.change_border_diagonal(weight)
      else raise 'invalid direction'
      end
    }
  end

  def add_cell_style(row,column)
    xf = @workbook.cell_xfs[@sheet_data[row][column].style_index]
    @workbook.fonts[xf.font_id].count += 1
    @workbook.fills[xf.fill_id].count += 1
    @workbook.borders[xf.border_id].count += 1
  end

  # finds first row which contains at least all strings in cells_content
  def find_first_row_with_content(cells_content)
    validate_workbook
    index = nil

    @sheet_data.each_with_index do |row, index|
      cells_content = cells_content.map { |header| header.to_s.downcase.strip }
      original_cells_content = row.map { |cell| cell.nil? ? '' : cell.value.to_s.downcase.strip }
      if (cells_content & original_cells_content).size == cells_content.size
        return index
      end
    end
    return nil
  end

  def validate_nonnegative(row_or_col)
    raise 'Row and Column arguments must be nonnegative' if row_or_col < 0
  end
  private :validate_nonnegative

end #end class
end
