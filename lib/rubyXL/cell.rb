module RubyXL
  class Cell < PrivateClass
    SHARED_STRING = 's'
    RAW_STRING = 'str'
    ERROR = 'e'

    attr_accessor :row, :column, :datatype, :style_index, :formula, :worksheet
    attr_reader :workbook,:formula_attributes

    def initialize(worksheet, row, column, value = nil, formula = nil, datatype = SHARED_STRING, style_index = 0, fmla_attr = {})
      @worksheet = worksheet

      @workbook = worksheet.workbook
      @row = row
      @column = column
      @datatype = datatype
      @value = value
      @formula=formula
      @style_index = style_index
      @formula_attributes = fmla_attr
    end

    def value(args = {})
      raw_values = args.delete(:raw) || false
      return @value if raw_values
      return @workbook.num_to_date(@value) if is_date?
      @value
    end

    def is_date?
      return false if @value.is_a?(String)
      tmp_num_fmt = @workbook.num_fmts_by_id[Integer(xf_id()[:numFmtId])]
      num_fmt = tmp_num_fmt && tmp_num_fmt.format_code
      num_fmt && workbook.date_num_fmt?(num_fmt)
    end

    # changes fill color of cell
    def change_fill(rgb='ffffff')
      validate_worksheet
      Color.validate_color(rgb)
      @style_index = modify_fill(@workbook, @style_index,rgb)
    end

    # Changes font name of cell
    def change_font_name(font_name = 'Verdana')
      validate_worksheet
      # Get copy of font object with modified name
      font = workbook.fonts[font_id()].dup
      font.set_name(font_name)
      # Update font and xf array
      change_font(font)
    end

    # Changes font size of cell
    def change_font_size(font_size=10)
      validate_worksheet
      if font_size.is_a?(Integer) || font_size.is_a?(Float)
        # Get copy of font object with modified size
        font = workbook.fonts[font_id()].dup
        font.set_size(font_size)
        # Update font and xf array
        change_font(font)
      else
        raise 'Argument must be a number'
      end
    end

    # Changes font color of cell
    def change_font_color(font_color='000000')
      validate_worksheet
      #if arg is a color name, convert to integer
      Color.validate_color(font_color)
      # Get copy of font object with modified color
      font = workbook.fonts[font_id()].dup
      font.set_rgb_color(font_color)
      # Update font and xf array
      change_font(font)
    end

    # Changes font italics settings of cell
    def change_font_italics(italicized=false)
      validate_worksheet
      # Get copy of font object with modified italics settings
      font = workbook.fonts[font_id()].dup
      font.set_italic(italicized)
      # Update font and xf array
      change_font(font)
    end

    # Changes font bold settings of cell
    def change_font_bold(bolded=false)
      validate_worksheet
      # Get copy of font object with modified bold settings
      font = workbook.fonts[font_id()].dup
      font.set_bold(bolded)
      # Update font and xf array
      change_font(font)
    end

    # Changes font underline settings of cell
    def change_font_underline(underlined=false)
      validate_worksheet
      # Get copy of font object with modified underline settings
      font = workbook.fonts[font_id()].dup
      font.set_underline(underlined)
      # Update font and xf array
      change_font(font)
    end

    # Changes font strikethrough settings of cell
    def change_font_strikethrough(struckthrough=false)
      validate_worksheet
      # Get copy of font object with modified strikethrough settings
      font = workbook.fonts[font_id()].dup
      font.set_strikethrough(struckthrough)
      # Update font and xf array
      change_font(font)
    end

    # Helper method to update the font array and xf array
    def change_font(font)
      # Modify font array and retrieve new font id
      new_font_id = modify_font(@workbook, font, font_id())
      # Get copy of xf object with modified font id
      xf = deep_copy(xf_id())
      xf[:fontId] = new_font_id
      # Modify xf array and retrieve new xf id
      @style_index = modify_xf(@workbook, xf)
    end

    # changes horizontal alignment of cell
    def change_horizontal_alignment(alignment='center')
      validate_worksheet
      validate_horizontal_alignment(alignment)
      @style_index = modify_alignment(@workbook,@style_index,true,alignment)
    end

    # changes vertical alignment of cell
    def change_vertical_alignment(alignment='center')
      validate_worksheet
      validate_vertical_alignment(alignment)
      @style_index = modify_alignment(@workbook,@style_index,false,alignment)
    end

    # changes wrap of cell
    def change_text_wrap(wrap=false)
      validate_worksheet
      validate_text_wrap(wrap)
      @style_index = modify_text_wrap(@workbook,@style_index,wrap)
    end

    # changes top border of cell
    def change_border_top(weight='thin')
      change_border(:top, weight)
    end

    # changes left border of cell
    def change_border_left(weight='thin')
      change_border(:left, weight)
    end

    # changes right border of cell
    def change_border_right(weight='thin')
      change_border(:right, weight)
    end

    # changes bottom border of cell
    def change_border_bottom(weight='thin')
      change_border(:bottom, weight)
    end

    # changes diagonal border of cell
    def change_border_diagonal(weight='thin')
      change_border(:diagonal, weight)
    end

    # changes contents of cell, with formula option
    def change_contents(data, formula=nil)
      validate_worksheet
      @datatype = RAW_STRING

      if data.is_a?(Date) || data.is_a?(DateTime)
        data = @workbook.date_to_num(data)
      end

      if (data.is_a?Integer) || (data.is_a?Float)
        @datatype = ''
      end

      @value=data
      @formula=formula
    end

    # returns if font is italicized
    def is_italicized()
      validate_worksheet
      @workbook.fonts[font_id()].is_italic
    end

    # returns if font is bolded
    def is_bolded()
      validate_worksheet
      @workbook.fonts[font_id()].is_bold
    end

    # returns if font is underlined
    def is_underlined()
      validate_worksheet
      xf = @workbook.get_style_attributes(@workbook.get_style(@style_index))
      @workbook.fonts[font_id()].is_underlined
    end

    # returns if font has a strike through it
    def is_struckthrough()
      validate_worksheet
      xf = @workbook.get_style_attributes(@workbook.get_style(@style_index))
      @workbook.fonts[font_id()].is_strikethrough
    end

    # returns cell's font name
    def font_name()
      validate_worksheet
      @workbook.fonts[font_id()].get_name
    end

    # returns cell's font size
    def font_size()
      validate_worksheet
      return @workbook.fonts[font_id()].get_size
    end

    # returns cell's font color
    def font_color()
      validate_worksheet
      @workbook.fonts[font_id()].get_rgb_color || '000000' #black
    end

    # returns cell's fill color
    def fill_color()
      validate_worksheet
      xf = @workbook.get_style_attributes(@workbook.get_style(@style_index))
      return @workbook.get_fill_color(xf)
    end

    # returns cell's horizontal alignment
    def horizontal_alignment()
      validate_worksheet
      xf_obj = @workbook.get_style(@style_index)
      if xf_obj[:alignment].nil? || xf_obj[:alignment][:attributes].nil?
        return nil
      end
      xf_obj[:alignment][:attributes][:horizontal].to_s
    end

    # returns cell's vertical alignment
    def vertical_alignment()
      validate_worksheet
      xf_obj = @workbook.get_style(@style_index)
      if xf_obj[:alignment].nil? || xf_obj[:alignment][:attributes].nil?
        return nil
      end
      xf_obj[:alignment][:attributes][:vertical].to_s
    end

    # returns cell's wrap
    def text_wrap()
      validate_worksheet
      xf_obj = @workbook.get_style(@style_index)
      if xf_obj[:alignment].nil? || xf_obj[:alignment][:attributes].nil?
        return nil
      end
      (xf_obj[:alignment][:attributes][:wrapText]== "1")
    end

    # returns cell's top border
    def border_top()
      return get_border(:top)
    end

    # returns cell's left border
    def border_left()
      return get_border(:left)
    end

    # returns cell's right border
    def border_right()
      return get_border(:right)
    end

    # returns cell's bottom border
    def border_bottom()
      return get_border(:bottom)
    end

    # returns cell's diagonal border
    def border_diagonal()
      return get_border(:diagonal)
    end

    def inspect
      str = "(#{@row},#{@column}): #{@value}" 
      str += " =#{@formula}" if @formula
      str += ", datatype = #{@datatype}, style_index = #{@style_index}"
      return str
    end

    private

    def change_border(direction, weight)
      validate_worksheet
      validate_border(weight)
      @style_index = modify_border(@workbook, @style_index)
      border = @workbook.borders[xf_id()[:borderId]]
      border.edges[direction.to_s] ||= RubyXL::BorderEdge.new
      border.edges[direction.to_s].style = weight
    end

    def get_border(direction)
      validate_worksheet

      border = @workbook.borders[xf_id()[:borderId]]
      edge = border.edges[direction.to_s]
      edge && edge.style
    end

    def validate_workbook()
      unless @workbook.nil? || @workbook.worksheets.nil?
        @workbook.worksheets.each do |sheet|
          unless sheet.nil? || sheet.sheet_data.nil? || sheet.sheet_data[@row].nil?
            if sheet.sheet_data[@row][@column] == self
              return
            end
          end
        end
      end
      raise "This cell #{self} is not in workbook #{@workbook}"
    end

    def validate_worksheet()
      return if @worksheet && @worksheet[@row][@column] == self
      raise "This cell #{self} is not in worksheet #{worksheet}"
    end

    def xf_id()
      @workbook.get_style_attributes(@workbook.get_style(@style_index.to_s))
    end

    def font_id()
      xf_id()[:fontId]
    end
  end
end
