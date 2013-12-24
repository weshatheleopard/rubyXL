module RubyXL
class Worksheet < PrivateClass
  include Enumerable

  attr_accessor :sheet_name, :sheet_id, :sheet_data, :column_ranges, :merged_cells, :pane,
                :validations, :sheet_view, :legacy_drawing, :extLst, :workbook,
                :row_styles, :drawings

  SHEET_NAME_TEMPLATE = 'Sheet%d'

  def initialize(workbook, sheet_name = nil, sheet_data= [[nil]], cols=[], merged_cells=[])
    @workbook = workbook

    @sheet_name = sheet_name || get_default_name
    @sheet_id = nil
    @sheet_data = sheet_data
    @column_ranges = cols
    @merged_cells = merged_cells || []
    @row_styles={}
    @sheet_view = {
                    :attributes => {
                                      :workbookViewId => 0, :zoomScale => 100, :tabSelected => 1, :view=>'normalLayout', :zoomScaleNormal => 100
                                   }
                  }
    @extLst = nil
    @legacy_drawing = nil
    @drawings = []
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
  def [](row=0)
    return @sheet_data[row]
  end

  def each
    @sheet_data.each {|i| yield i}
  end

  #returns 2d array of just the cell values (without style or formula information)
  def extract_data(args = {})
    raw_values = args.delete(:raw) || false
    return @sheet_data.map {|row| row.map {|c| if c.is_a?(Cell) then c.value(:raw => raw_values) else nil end}}
  end

  def get_table(headers=[],opts={})
    validate_workbook

    if !headers.is_a?(Array)
      headers = [headers]
    end

    row_num = find_first_row_with_content(headers)
  	
    return nil if row_num.nil?

  	table_hash = {}
  	table_hash[:table] = []

    header_row = @sheet_data[row_num]
  	header_row.each_with_index do |header_cell, index|
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
  	  cell_test= (!cell.nil? && !cell.value.nil?)

      while cell_test || (table_hash[:table][table_index] && !table_hash[:table][table_index].empty?)

  		  table_hash[header] << cell.value if cell_test

    		table_index = current_row - original_row

                if cell_test then
                  table_hash[:table][table_index] ||= {}
                  table_hash[:table][table_index][header] = cell.value 
                end

    		current_row += 1
    		if @sheet_data[current_row].nil?
    		  cell = nil
  		  else
      		cell = @sheet_data[current_row][index]
    		end
        cell_test= (!cell.nil? && !cell.value.nil?)
  	  end
	  end

	  return table_hash
  end

  #changes color of fill in (zer0 indexed) row
  def change_row_fill(row = 0, rgb = 'ffffff')
    validate_workbook
    validate_nonnegative(row)
    ensure_cell_exists(row)
    Color.validate_color(rgb)

    if @row_styles[(Integer(row)+1).to_s].nil?
      @row_styles[(Integer(row)+1).to_s] = {}
      @row_styles[(Integer(row)+1).to_s][:style] = '0'
    end

    @row_styles[(Integer(row)+1).to_s][:style] = modify_fill(@workbook,Integer(@row_styles[(Integer(row)+1).to_s][:style]),rgb)

    @sheet_data[Integer(row)].each do |c|
      unless c.nil?
        c.change_fill(rgb)
      end
    end
  end

  # Changes font name of row
  def change_row_font_name(row=0, font_name='Verdana')
    # Get style object
    xf_id = xf_id(get_row_style(row))
    # Get copy of font object with modified name
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font[:name][:attributes][:val] = font_name.to_s
    # Update font and xf array
    change_row_font(row, Worksheet::NAME, font_name, font, xf_id)
  end

  # Changes font size of row
  def change_row_font_size(row=0, font_size=10)
    # Get style object
    xf_id = xf_id(get_row_style(row))
    # Get copy of font object with modified size
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font[:sz][:attributes][:val] = font_size
    # Update font and xf array
    change_row_font(row, Worksheet::SIZE, font_size, font, xf_id)
  end

  # Changes font color of row
  def change_row_font_color(row=0, font_color='000000')
    Color.validate_color(font_color)
    # Get style object
    xf_id = xf_id(get_row_style(row))
    # Get copy of font object with modified color
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font = modify_font_color(font, font_color.to_s)
    # Update font and xf array
    change_row_font(row, Worksheet::COLOR, font_color, font, xf_id)
  end

  # Changes font italics settings of row
  def change_row_italics(row=0, italicized=false)
    # Get style object
    xf_id = xf_id(get_row_style(row))
    # Get copy of font object with modified italics settings
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font = modify_font_italics(font, italicized)
    # Update font and xf array
    change_row_font(row, Worksheet::ITALICS, italicized, font, xf_id)
  end

  # Changes font bold settings of row
  def change_row_bold(row=0, bolded=false)
    # Get style object
    xf_id = xf_id(get_row_style(row))
    # Get copy of font object with modified bold settings
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font = modify_font_bold(font, bolded)
    # Update font and xf array
    change_row_font(row, Worksheet::BOLD, bolded, font, xf_id)
  end

  # Changes font underline settings of row
  def change_row_underline(row=0, underlined=false)
    # Get style object
    xf_id = xf_id(get_row_style(row))
    # Get copy of font object with modified underline settings
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font = modify_font_underline(font, underlined)
    # Update font and xf array
    change_row_font(row, Worksheet::UNDERLINE, underlined, font, xf_id)
  end

  # Changes font strikethrough settings of row
  def change_row_strikethrough(row=0, struckthrough=false)
    # Get style object
    xf_id = xf_id(get_row_style(row))
    # Get copy of font object with modified strikethrough settings
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font = modify_font_strikethrough(font, struckthrough)
    # Update font and xf array
    change_row_font(row, Worksheet::STRIKETHROUGH, struckthrough, font, xf_id)
  end

  def change_row_height(row=0, height=10)
    validate_workbook
    validate_nonnegative(row)

    ensure_cell_exists(row)

    if height.to_i.to_s == height.to_s
      height = Integer(height)
    elsif height.to_f.to_s == height.to_s
      height = Float(height)
    else
      raise 'You must enter a number for the height'
    end

    if @row_styles[(row+1).to_s].nil?
      @row_styles[(row+1).to_s] = {}
      @row_styles[(row+1).to_s][:style] = '0'
    end
    @row_styles[(row+1).to_s][:height] = height
    @row_styles[(row+1).to_s][:customHeight] = '1'
  end

  def change_row_horizontal_alignment(row=0,alignment='center')
    validate_workbook
    validate_nonnegative(row)
    validate_horizontal_alignment(alignment)
    change_row_alignment(row,alignment,true)
  end

  def change_row_vertical_alignment(row=0,alignment='center')
    validate_workbook
    validate_nonnegative(row)
    validate_vertical_alignment(alignment)
    change_row_alignment(row,alignment,false)
  end

  def change_row_border_top(row=0,weight='thin')
    change_row_border(row, :top, weight)
  end

  def change_row_border_left(row=0,weight='thin')
    change_row_border(row, :left, weight)
  end

  def change_row_border_right(row=0,weight='thin')
    change_row_border(row, :right, weight)
  end

  def change_row_border_bottom(row=0,weight='thin')
    change_row_border(row, :bottom, weight)
  end

  def change_row_border_diagonal(row=0,weight='thin')
    change_row_border(row, :diagonal, weight)
  end

  # Changes font name of column
  def change_column_font_name(col=0, font_name='Verdana')
    # Get style object
    xf_id = xf_id(get_col_style(col))
    # Get copy of font object with modified name
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font[:name][:attributes][:val] = font_name.to_s
    # Update font and xf array
    change_column_font(col, Worksheet::NAME, font_name, font, xf_id)
  end

  # Changes font size of column
  def change_column_font_size(col=0, font_size=10)
    # Get style object
    xf_id = xf_id(get_col_style(col))
    # Get copy of font object with modified size
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font[:sz][:attributes][:val] = font_size
    # Update font and xf array
    change_column_font(col, Worksheet::SIZE, font_size, font, xf_id)
  end

  # Changes font color of column
  def change_column_font_color(col=0, font_color='000000')
    Color.validate_color(font_color)
    # Get style object
    xf_id = xf_id(get_col_style(col))
    # Get copy of font object with modified color
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font = modify_font_color(font, font_color.to_s)
    # Update font and xf array
    change_column_font(col, Worksheet::COLOR, font_color, font, xf_id)
  end

  # Changes font italics settings of column
  def change_column_italics(col=0, italicized=false)
    # Get style object
    xf_id = xf_id(get_col_style(col))
    # Get copy of font object with modified italics settings
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font = modify_font_italics(font, italicized)
    # Update font and xf array
    change_column_font(col, Worksheet::ITALICS, italicized, font, xf_id)
  end

  # Changes font bold settings of column
  def change_column_bold(col=0, bolded=false)
    # Get style object
    xf_id = xf_id(get_col_style(col))
    # Get copy of font object with modified bold settings
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font = modify_font_bold(font, bolded)
    # Update font and xf array
    change_column_font(col, Worksheet::BOLD, bolded, font, xf_id)
  end

  # Changes font underline settings of column
  def change_column_underline(col=0, underlined=false)
    # Get style object
    xf_id = xf_id(get_col_style(col))
    # Get copy of font object with modified underline settings
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font = modify_font_underline(font, underlined)
    # Update font and xf array
    change_column_font(col, Worksheet::UNDERLINE, underlined, font, xf_id)
  end

  # Changes font strikethrough settings of column
  def change_column_strikethrough(col=0, struckthrough=false)
    # Get style object
    xf_id = xf_id(get_col_style(col))
    # Get copy of font object with modified strikethrough settings
    font = deep_copy(@workbook.fonts[xf_id[:fontId].to_s][:font])
    font = modify_font_strikethrough(font, struckthrough)
    # Update font and xf array
    change_column_font(col, Worksheet::STRIKETHROUGH, struckthrough, font, xf_id)
  end

  def change_column_width(col = 0, width = 13)
    validate_workbook
    validate_nonnegative(col)
    ensure_cell_exists(0, col)

    RubyXL::ColumnRange.update(col, @column_ranges, { 'width' => width, 'customWidth' => 1 })
  end

  def get_column_style_index(col)
    range = RubyXL::ColumnRange.find(col, @column_ranges)
    (range && range.style_index) || 0
  end

  def change_column_fill(col=0, color_index='ffffff')
    validate_workbook
    validate_nonnegative(col)
    Color.validate_color(color_index)
    ensure_cell_exists(0, col)

    new_style_index = modify_fill(@workbook, get_column_style_index(col), color_index)
    RubyXL::ColumnRange.update(col, @column_ranges, { 'style' => new_style_index })

    @sheet_data.each { |row|
      c = row[col]
      c.change_fill(color_index) if c
    }
  end

  def change_column_horizontal_alignment(col=0,alignment='center')
    validate_horizontal_alignment(alignment)
    change_column_alignment(col,alignment,true)
  end

  def change_column_vertical_alignment(col=0,alignment='center')
    validate_vertical_alignment(alignment)
    change_column_alignment(col,alignment,false)
  end

  def change_column_border_top(col=0,weight='thin')
    change_column_border(col,:top,weight)
  end

  def change_column_border_left(col=0,weight='thin')
    change_column_border(col,:left,weight)
  end

  def change_column_border_right(col=0,weight='thin')
    change_column_border(col,:right,weight)
  end

  def change_column_border_bottom(col=0,weight='thin')
    change_column_border(col,:bottom,weight)
  end

  def change_column_border_diagonal(col=0,weight='thin')
    change_column_border(col,:diagonal,weight)
  end

  # merges cells within a rectangular range
  def merge_cells(row1 = 0, col1 = 0, row2 = 0, col2 = 0)
    validate_workbook
    @merged_cells << "#{Cell.ind2ref(row1, col1)}:#{Cell.ind2ref(row2, col2)}"
  end

  def add_cell(row=0, column=0, data='', formula=nil,overwrite=true)
    validate_workbook
    validate_nonnegative(row)
    validate_nonnegative(column)
    ensure_cell_exists(row, column)

    datatype = (formula.nil?) ? RubyXL::Cell::RAW_STRING : ''

    if overwrite || @sheet_data[row][column].nil?
      @sheet_data[row][column] = Cell.new(self,row,column,data,formula,datatype)

      if (data.is_a?Integer) || (data.is_a?Float)
        @sheet_data[row][column].datatype = ''
      end
      col = RubyXL::ColumnRange.find(column, @column_ranges)

      if @row_styles[(row+1).to_s] != nil
        @sheet_data[row][column].style_index = @row_styles[(row+1).to_s][:style]
      elsif col != nil
        @sheet_data[row][column].style_index = col.style_index
      end
    end

    add_cell_style(row,column)

    return @sheet_data[row][column]
  end

  def add_cell_obj(cell, overwrite=true)
    validate_workbook

    if cell.nil?
      return cell
    end

    row = cell.row
    column = cell.column

    validate_nonnegative(row)
    validate_nonnegative(column)
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

    row_num.upto(@sheet_data.size) do |index|
      @row_styles[(index-1).to_s] = deep_copy(@row_styles[index.to_s])
    end
    @row_styles.delete(@sheet_data.size.to_s)

    #change row styles
    # raise row_styles.inspect

    #change cell row numbers
    (row_index...(@sheet_data.size-1)).each do |index|
      @sheet_data[index].map {|c| c.row -= 1 if c}
    end

    return deleted
  end

  #inserts row at row_index, pushes down, copies style from below (row previously at that index)
  #USE OF THIS METHOD will break formulas which reference cells which are being "pushed down"
  def insert_row(row_index=0)
    validate_workbook
    validate_nonnegative(row_index)

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

        if @row_styles[(row_num+1).to_s].nil?
          @row_styles[(row_num+1).to_s] = {:style=>0}
        end
        if old_cell.style_index != 0 && old_cell.style_index.to_s != @row_styles[(row_num+1).to_s][:style].to_s
          c = Cell.new(self,row_index,i)
          c.style_index = old_cell.style_index
          @sheet_data[row_index][i] = c
        end
      end
    end

    #copy row styles from row above, (or below if first row)
    (@row_styles.size+1).downto(row_num+1) do |i|
      @row_styles[i.to_s] = @row_styles[(i-1).to_s]
    end
    if row_index > 0
      @row_styles[row_num.to_s] = @row_styles[(row_num-1).to_s]
    else
      @row_styles[row_num.to_s] = nil#@row_styles[(row_num+1).to_s]
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
    validate_nonnegative(col_index)
    ensure_cell_exists(0, col_index)

    old_range = col_index > 0 ? RubyXL::ColumnRange.find(col_index, @column_ranges) : RubyXL::ColumnRange.new

    #go through each cell in column
    @sheet_data.each_with_index do |row, row_index|
      old_cell = row[col_index]
      new_cell = nil

      if old_cell && old_cell.style_index != 0 &&
           old_range && old_range.style_index != old_cell.style_index.to_i then
        new_cell = Cell.new(self, row_index, col_index)
        new_cell.style_index = old_cell.style_index
      end

      row.insert(col_index, new_cell)
    end

    ColumnRange.insert_column(col_index, @column_ranges)

    #update column numbers
    @sheet_data.each { |row|
      (col_index + 1).upto(row.size) { |col|
        row[col].column = col unless row[col].nil?
      }
    }
  end

  def insert_cell(row=0,col=0,data=nil,formula=nil,shift=nil)
    validate_workbook
    validate_nonnegative(row)
    validate_nonnegative(col)
    ensure_cell_exists(row, col)

    if shift && shift != :right && shift != :down
      raise 'invalid shift option'
    end

    if shift == :right
      @sheet_data[row].insert(col,nil)
      (row...(@sheet_data[row].size)).each do |index|
        if @sheet_data[row][index].is_a?(Cell)
          @sheet_data[row][index].column += 1
        end
      end
    elsif shift == :down
      @sheet_data << Array.new(@sheet_data[row].size)
      (@sheet_data.size-1).downto(row+1) do |index|
        @sheet_data[index][col] = @sheet_data[index-1][col]
      end
    end

    return add_cell(row,col,data,formula)
  end

  # by default, only sets cell to nil
  # if :left is specified, method will shift row contents to the right of the deleted cell to the left
  # if :up is specified, method will shift column contents below the deleted cell upward
  def delete_cell(row=0,col=0,shift=nil)
    validate_workbook
    validate_nonnegative(row)
    validate_nonnegative(col)
    if @sheet_data.size <= row || @sheet_data[row].size <= col
      return nil
    end

    cell = @sheet_data[row][col]
    @sheet_data[row][col]=nil

    if shift && shift != :left && shift != :up
      raise 'invalid shift option'
    end

    if shift == :left
      @sheet_data[row].delete_at(col)
      @sheet_data[row] << nil
      (col...(@sheet_data[row].size)).each do |index|
        if @sheet_data[row][index].is_a?(Cell)
          @sheet_data[row][index].column -= 1
        end
      end
    elsif shift == :up
      (row...(@sheet_data.size-1)).each do |index|
        @sheet_data[index][col] = @sheet_data[index+1][col]
        if @sheet_data[index][col].is_a?(Cell)
          @sheet_data[index][col].row -= 1
        end
      end
      if @sheet_data.last[col].is_a?(Cell)
        @sheet_data.last[col].row -= 1
      end
    end

    return cell
  end

  def get_row_fill(row=0)
    validate_workbook
    validate_nonnegative(row)

    if @sheet_data.size <= row
      return nil
    end

    if @row_styles[(row+1).to_s].nil?
      return "ffffff" #default, white
    end

    xf = xf_attr_row(row)

    return @workbook.get_fill_color(xf)
  end

  def get_row_font_name(row=0)
    validate_workbook
    validate_nonnegative(row)

    if @sheet_data.size <= row
      return nil
    end

    if @row_styles[(row+1).to_s].nil?
      return 'Verdana'
    end

    xf = xf_attr_row(row)

    return @workbook.fonts[xf[:fontId].to_s][:font][:name][:attributes][:val]
  end

  def get_row_font_size(row=0)
    validate_workbook
    validate_nonnegative(row)

    if @sheet_data.size <= row
      return nil
    end

    if @row_styles[(row+1).to_s].nil?
      return '10'
    end

    xf = xf_attr_row(row)

    return @workbook.fonts[xf[:fontId].to_s][:font][:sz][:attributes][:val]
  end

  def get_row_font_color(row=0)
    validate_workbook
    validate_nonnegative(row)

    if @sheet_data.size <= row
      return nil
    end

    if @row_styles[(row+1).to_s].nil?
      return '000000'
    end

    xf = xf_attr_row(row)

    color = @workbook.fonts[xf[:fontId].to_s][:font][:color]

    if color.nil? || color[:attributes].nil? || color[:attributes][:rgb].nil?
      return '000000'
    end

    return color[:attributes][:rgb]
  end

  def is_row_italicized(row=0)
    return get_row_bool(row,:i)
  end

  def is_row_bolded(row=0)
    return get_row_bool(row,:b)
  end

  def is_row_underlined(row=0)
    return get_row_bool(row,:u)
  end

  def is_row_struckthrough(row=0)
    return get_row_bool(row,:strike)
  end


  def get_row_height(row=0)
    validate_workbook
    validate_nonnegative(row)

    if @sheet_data.size <= row
      return nil
    end

    if @row_styles[(row+1).to_s].nil?
      return 13
    else
      @row_styles[(row+1).to_s][:height]
    end
  end

  def get_row_horizontal_alignment(row=0)
    return get_row_alignment(row,true)
  end

  def get_row_vertical_alignment(row=0)
    return get_row_alignment(row,false)
  end

  def get_row_border_top(row=0)
    return get_row_border(row,:top)
  end

  def get_row_border_left(row=0)
    return get_row_border(row,:left)
  end

  def get_row_border_right(row=0)
    return get_row_border(row,:right)
  end

  def get_row_border_bottom(row=0)
    return get_row_border(row,:bottom)
  end

  def get_row_border_diagonal(row=0)
    return get_row_border(row,:diagonal)
  end

  def get_column_font_name(col=0)
    validate_workbook
    validate_nonnegative(col)

    if @sheet_data[0].size <= col
      return nil
    end

    style_index = get_cols_style_index(col)

    return @workbook.fonts[font_id( style_index ).to_s][:font][:name][:attributes][:val]
  end

  def get_column_font_size(col=0)
    validate_workbook
    validate_nonnegative(col)

    if @sheet_data[0].size <= col
      return nil
    end

    style_index = get_cols_style_index(col)

    return @workbook.fonts[font_id( style_index ).to_s][:font][:sz][:attributes][:val]
  end

  def get_column_font_color(col=0)
    validate_workbook
    validate_nonnegative(col)

    if @sheet_data[0].size <= col
      return nil
    end

    style_index = get_cols_style_index(col)

    font = @workbook.fonts[font_id( style_index ).to_s][:font]

    if font[:color].nil? || font[:color][:attributes].nil? || font[:color][:attributes][:rgb].nil?
      return '000000'
    end

    return font[:color][:attributes][:rgb]

  end

  def is_column_italicized(col=0)
    get_column_bool(col,:i)
  end

  def is_column_bolded(col=0)
    get_column_bool(col,:b)
  end

  def is_column_underlined(col=0)
    get_column_bool(col,:u)
  end

  def is_column_struckthrough(col=0)
    get_column_bool(col,:strike)
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

    if @sheet_data[0].size <= col
      return nil
    end

    style_index = get_cols_style_index(col)

    if style_index == 0
      return "ffffff" #default, white
    end

    return @workbook.get_fill_color(@workbook.cell_xfs[:xf][style_index][:attributes])
  end

  def get_column_horizontal_alignment(col=0)
    get_column_alignment(col, :horizontal)
  end

  def get_column_vertical_alignment(col=0)
    get_column_alignment(col, :vertical)
  end

  def get_column_border_top(col=0)
    return get_column_border(col,:top)
  end

  def get_column_border_left(col=0)
    return get_column_border(col,:left)
  end

  def get_column_border_right(col=0)
    return get_column_border(col,:right)
  end

  def get_column_border_bottom(col=0)
    return get_column_border(col,:bottom)
  end

  def get_column_border_diagonal(col=0)
    return get_column_border(col,:diagonal)
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
  def xf_attr_row(row)
    row_style = @row_styles[(row+1).to_s][:style]
    return @workbook.get_style_attributes(@workbook.get_style(row_style))
  end

  def xf_attr_col(column)
    @workbook.get_style_attributes(@workbook.get_style(RubyXL::ColumnRange.find(column).style_index))
  end

  def get_row_bool(row,property)
    validate_workbook
    validate_nonnegative(row)

    if @sheet_data.size <= row
      return nil
    end

    if @row_styles[(row+1).to_s].nil?
      return false
    end

    xf = xf_attr_row(row)

    return !@workbook.fonts[xf[:fontId].to_s][:font][property].nil?
  end

  def get_row_alignment(row,is_horizontal)
    validate_workbook
    validate_nonnegative(row)

    if @sheet_data.size <= row || @row_styles[(row+1).to_s].nil?
      return nil
    end

    xf_obj = @workbook.get_style(@row_styles[(row+1).to_s][:style])

    if xf_obj[:alignment].nil? || xf_obj[:alignment][:attributes].nil?
      return nil
    end

    if is_horizontal
      return xf_obj[:alignment][:attributes][:horizontal].to_s
    else
      return xf_obj[:alignment][:attributes][:vertical].to_s
    end
  end

  def get_row_border(row, border_direction)
    validate_workbook
    validate_nonnegative(row)

    return nil if @sheet_data.size <= row || @row_styles[(row+1).to_s].nil?

    border = @workbook.borders[xf_attr_row(row)[:borderId]]
    edge = border && border.edges[border_direction.to_s]
    edge && edge.style
  end

  def get_column_bool(col,property)
    validate_workbook
    validate_nonnegative(col)

    if @sheet_data[0].size <= col
      return nil
    end

    style_index = get_cols_style_index(col)

    return !@workbook.fonts[font_id( style_index ).to_s][:font][property].nil?
  end

  def get_column_alignment(col, type)
    validate_workbook
    validate_nonnegative(col)

    if @sheet_data[0].size <= col
      return nil
    end

    style_index = get_cols_style_index(col)

    xf_obj = @workbook.get_style(style_index)
    if xf_obj[:alignment].nil?
      return nil
    end

    return xf_obj[:alignment][:attributes][type]
  end

  def get_column_border(col, border_direction)
    validate_workbook
    validate_nonnegative(col)

    return nil if @sheet_data[0].size <= col

    xf = @workbook.get_style_attributes(@workbook.get_style(get_cols_style_index(col)))

    border = @workbook.borders[xf[:borderId]]
    edge = border && border.edges[border_direction.to_s]
    edge && edge.style
  end

  def deep_copy(hash)
    Marshal.load(Marshal.dump(hash))
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
    (range && range.style_index) || 0
  end

  # Helper method to update the row styles array
  # change_type - NAME or SIZE or COLOR etc
  # main method to change font, called from each separate font mutator method
  def change_row_font(row, change_type, arg, font, xf_id)
    validate_workbook
    validate_nonnegative(row)
    ensure_cell_exists(row)

    # Modify font array and retrieve new font id
    font_id = modify_font(@workbook, font, xf_id[:fontId].to_s)
    # Get copy of xf object with modified font id
    xf = deep_copy(xf_id)
    xf[:fontId] = Integer(font_id)
    # Modify xf array and retrieve new xf id
    @row_styles[(row+1).to_s][:style] = modify_xf(@workbook, xf)

    if @sheet_data[row].nil?
      @sheet_data[row] = []
    end

    @sheet_data[Integer(row)].each do |c|
      unless c.nil?
        font_switch(c, change_type, arg)
      end
    end
  end

  # Helper method to update the fonts and cell styles array
  # main method to change font, called from each separate font mutator method
  def change_column_font(col, change_type, arg, font, xf_id)
    validate_workbook
    validate_nonnegative(col)
    ensure_cell_exists(0, col)

    # Modify font array and retrieve new font id
    font_id = modify_font(@workbook, font, xf_id[:fontId].to_s)
    # Get copy of xf object with modified font id
    xf = deep_copy(xf_id)
    xf[:fontId] = Integer(font_id)
    # Modify xf array and retrieve new xf id
    modify_xf(@workbook, xf)

#    new_style_index = modify_fill(@workbook, get_column_style_index(col), color_index)
#    RubyXL::ColumnRange.update(col, @column_ranges, { 'style' => new_style_index })

    @sheet_data.each do |row|
      c = row[col]
      unless c.nil?
        font_switch(c, change_type, arg)
      end
    end
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

  # Helper method to get the font id for a style index
  def font_id(style_index)
    xf_id(style_index)[:fontId]
  end

  # Helper method to get the style attributes for a style index
  def xf_id(style_index)
    @workbook.get_style_attributes(@workbook.get_style(style_index))
  end

  # Helper method to get the style index for a row
  def get_row_style(row)
    if @row_styles[(row+1).to_s].nil?
      @row_styles[(row+1).to_s] = {}
      @row_styles[(row+1).to_s][:style] = '0'
      @workbook.fonts['0'][:count] += 1
    end
    return @row_styles[(row+1).to_s][:style]
  end

  # Helper method to get the style index for a column
  def get_col_style(col)
    range = RubyXL::ColumnRange.find(col, @column_ranges)

    return range.style_index if range
    
    @workbook.fonts['0'][:count] += 1

    0
  end

  def change_row_alignment(row,alignment,is_horizontal)
    validate_workbook
    validate_nonnegative(row)
    ensure_cell_exists(row)

    if @row_styles[(row+1).to_s].nil?
      @row_styles[(row+1).to_s] = {}
      @row_styles[(row+1).to_s][:style] = '0'
    end

    @row_styles[(row+1).to_s][:style] =
      modify_alignment(@workbook,@row_styles[(row+1).to_s][:style],is_horizontal,alignment)

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
    validate_nonnegative(col)
    ensure_cell_exists(0, col)

    new_style_index = modify_alignment(@workbook, get_column_style_index(col), is_horizontal, alignment)
    RubyXL::ColumnRange.update(col, @column_ranges, { 'style' => new_style_index })

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
    validate_nonnegative(row)
    validate_border(weight)
    ensure_cell_exists(row)

    if @row_styles[(row+1).to_s].nil?
      @row_styles[(row+1).to_s]= {}
      @row_styles[(row+1).to_s][:style] = '0'
    end
    @row_styles[(row+1).to_s][:style] = modify_border(@workbook, @row_styles[(row+1).to_s][:style])

    border = @workbook.borders[xf_attr_row(row)[:borderId]]
    border.edges[direction.to_s] ||= RubyXL::BorderEdge.new
    border.edges[direction.to_s].style = weight

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
    validate_nonnegative(col)
    validate_border(weight)
    ensure_cell_exists(0, col)

    new_style_index = modify_border(@workbook, get_column_style_index(col))
    RubyXL::ColumnRange.update(col, @column_ranges, { 'style' => new_style_index })

    xf = @workbook.get_style_attributes(@workbook.get_style(new_style_index))

    border = @workbook.borders[xf[:borderId]]
    border.edges[direction.to_s] ||= RubyXL::BorderEdge.new
    border.edges[direction.to_s].style = weight

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
    xf = @workbook.get_style_attributes(@workbook.get_style(@sheet_data[row][column].style_index))
    @workbook.fonts[xf[:fontId].to_s][:count] += 1
    @workbook.fills[xf[:fillId]].count += 1
    @workbook.borders[xf[:borderId]].count += 1
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

end #end class
end
