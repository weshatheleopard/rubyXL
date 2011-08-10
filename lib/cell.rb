module RubyXL
  class Cell < PrivateClass

    attr_accessor :row, :column, :datatype, :style_index, :value, :formula, :worksheet
    attr_reader :workbook

    def initialize(worksheet,row,column,value=nil,formula=nil,datatype='s',style_index=0)
      @worksheet = worksheet

      @workbook = worksheet.workbook
      @row = row
      @column = column
      @datatype = datatype
      @value = value
      @formula=formula
      @style_index = style_index
    end

    # changes fill color of cell
    def change_fill(rgb='ffffff')
      validate_worksheet
      Color.validate_color(rgb)
      @style_index = modify_fill(@workbook, @style_index,rgb)
    end

    # changes font name of cell
    def change_font_name(font_name='Verdana')
      validate_worksheet
      @style_index = modify_font(@workbook,@style_index)
      @workbook.fonts[font_id()][:font][:name][:attributes][:val] = font_name.to_s
    end

    # changes font size of cell
    def change_font_size(font_size=10)
      validate_worksheet
      if font_size.is_a?(Integer) || font_size.is_a?(Float)
        @style_index = modify_font(@workbook, @style_index)
        @workbook.fonts[font_id()][:font][:sz][:attributes][:val] = font_size
      else
        raise 'Argument must be a number'
      end
    end

    # changes font color of cell
    def change_font_color(font_color='000000')
      validate_worksheet
      #if arg is a color name, convert to integer
      Color.validate_color(font_color)

      @style_index = modify_font(@workbook,@style_index)
      font_id = font_id()
      if @workbook.fonts[font_id][:font][:color].nil?
        @workbook.fonts[font_id][:font][:color] = {:attributes => {:rgb => ''}}
      end
      @workbook.fonts[font_id][:font][:color][:attributes][:rgb] = font_color.to_s
    end

    # changes if font is italicized or not
    def change_font_italics(italicized=false)
      validate_worksheet
      @style_index = modify_font(@workbook,@style_index)
      if italicized
        @workbook.fonts[font_id()][:font][:i] = {}
      else
        @workbook.fonts[font_id()][:font][:i] = nil
      end
    end

    # changes if font is bolded or not
    def change_font_bold(bolded=false)
      validate_worksheet
      @style_index = modify_font(@workbook,@style_index)
      if bolded
        @workbook.fonts[font_id()][:font][:b] = {}
      else
        @workbook.fonts[font_id()][:font][:b] = nil
      end
    end

    # changes if font is underlined or not
    def change_font_underline(underlined=false)
      validate_worksheet
      @style_index = modify_font(@workbook,@style_index)

      if underlined
        @workbook.fonts[font_id()][:font][:u] = {}
      else
        @workbook.fonts[font_id()][:font][:u] = nil
      end
    end

    # changes if font has a strikethrough or not
    def change_font_strikethrough(struckthrough=false)
      validate_worksheet
      @style_index = modify_font(@workbook,@style_index)

      if struckthrough
        @workbook.fonts[font_id()][:font][:strike] = {}
      else
        @workbook.fonts[font_id()][:font][:strike] = nil
      end
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
      @datatype='str'
      if (data.is_a?Integer) || (data.is_a?Float)
        @datatype = ''
      end
      @value=data
      @formula=formula
    end

    # returns if font is italicized
    def is_italicized()
      validate_worksheet
      if @workbook.fonts[font_id()][:font][:i].nil?
        false
      else
        true
      end
    end

    # returns if font is bolded
    def is_bolded()
      validate_worksheet
      if @workbook.fonts[font_id()][:font][:b].nil?
        false
      else
        true
      end
    end

    # returns if font is underlined
    def is_underlined()
      validate_worksheet
      xf = @workbook.get_style_attributes(@workbook.get_style(@style_index))
      if @workbook.fonts[font_id()][:font][:u].nil?
        false
      else
        true
      end
    end

    # returns if font has a strike through it
    def is_struckthrough()
      validate_worksheet
      xf = @workbook.get_style_attributes(@workbook.get_style(@style_index))
      if @workbook.fonts[font_id()][:font][:strike].nil?
        false
      else
        true
      end
    end

    # returns cell's font name
    def font_name()
      validate_worksheet
      @workbook.fonts[font_id()][:font][:name][:attributes][:val]
    end

    # returns cell's font size
    def font_size()
      validate_worksheet
      return @workbook.fonts[font_id()][:font][:sz][:attributes][:val]
    end

    # returns cell's font color
    def font_color()
      validate_worksheet
      if @workbook.fonts[font_id()][:font][:color].nil?
        '000000' #black
      else
        @workbook.fonts[font_id()][:font][:color][:attributes][:rgb]
      end
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

    # returns Excel-style cell string from matrix indices
    def Cell.convert_to_cell(row=0,col=0)
      row_string = (row + 1).to_s #+1 for 0 indexing
      col_string = ''

      if row < 0 || col < 0
        raise 'Invalid input: cannot convert negative numbers'
      end

      unless col == 0
        col_length = 1+Integer(Math.log(col) / Math.log(26)) #opposite of 26**
      else
        col_length = 1
      end

      1.upto(col_length) do |i|

        #for the last digit, 0 should mean A. easy way to do this.
        if i == col_length
          col+=1
        end

        if col >= 26**(col_length-i)
          int_val = col / 26**(col_length-i) #+1 for 0 indexing
          int_val += 64 #converts 1 to A, etc.

          col_string += int_val.chr

          #intval multiplier decrements by placeholder, essentially
          #a B subtracts more than an A this way.
          col -= (int_val-64)*26**(col_length-i)
        end
      end
      col_string+row_string
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
      @style_index = modify_border(@workbook,@style_index)
      if @workbook.borders[border_id()][:border][direction][:attributes].nil?
        @workbook.borders[border_id()][:border][direction][:attributes] = { :style => nil }
      end
      @workbook.borders[border_id()][:border][direction][:attributes][:style] = weight.to_s
    end

    def get_border(direction)
      validate_worksheet

      if @workbook.borders[border_id()][:border][direction][:attributes].nil?
        return nil
      end
      return @workbook.borders[border_id()][:border][direction][:attributes][:style]
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
      if !@worksheet.nil? && @worksheet[@row][@column] == self
        return
      else
        raise "This cell #{self} is not in worksheet #{worksheet}"
      end
    end

    def xf_id()
      @workbook.get_style_attributes(@workbook.get_style(@style_index.to_s))
    end

    def border_id()
      xf_id()[:borderId].to_s
    end

    def font_id()
      xf_id()[:fontId].to_s
    end
  end
end
