module RubyXL
  class Cell < PrivateClass

    attr_accessor :row, :column, :datatype, :style_index, :formula, :worksheet
    attr_reader :workbook,:formula_attributes

    def initialize(worksheet,row,column,value=nil,formula=nil,datatype='s',style_index=0, fmla_attr={})
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
      
      if is_date?
        return @workbook.num_to_date(@value)
      else
        return @value
      end
    end

    def is_date?
      if !@value.is_a?(String)
        if @workbook.num_fmts_by_id
          num_fmt_id = xf_id()[:numFmtId]
          tmp_num_fmt = @workbook.num_fmts_by_id[num_fmt_id]
          num_fmt = (tmp_num_fmt &&tmp_num_fmt[:attributes] && tmp_num_fmt[:attributes][:formatCode]) ? tmp_num_fmt[:attributes][:formatCode] : nil
          if num_fmt && workbook.date_num_fmt?(num_fmt)
            return true
          end
        end
      end
      return false
    end

    def is_date_format?(num_fmt)
      skip_chars = ['$', '-', '+', '/', '(', ')', ':', ' ']
      num_chars = ['0', '#', '?']
      non_date_formats = ['0.00E+00', '##0.0E+0', 'General', 'GENERAL', 'general', '@']
      date_chars = ['y','m','d','h','s']

      state = 0
      s = ''
      num_fmt.split(//).each do |c|
        if state == 0
          if c == '"'
            state = 1
          elsif ['\\', '_', '*'].include?(c)
            state = 2
          elsif skip_chars.include?(c)
            next
          else
            s << c
          end
        elsif state == 1
          if c == '"'
            state = 0
          end
        elsif state == 2
          state = 0
        end
      end
      s.gsub!(/\[[^\]]*\]/, '')
      if non_date_formats.include?(s)
        return false
      end
      separator = ';'
      got_sep = 0
      date_count = 0
      num_count = 0
      s.split(//).each do |c|
        if date_chars.include?(c)
          date_count += 1
        elsif num_chars.include?(c)
          num_count += 1
        elsif c == separator
          got_sep = 1
        end
      end
      if date_count > 0 && num_count == 0
        return true
      elsif num_count > 0 && date_count == 0
        return false
      elsif date_count
        # ambiguous result
      elsif got_sep == 0
        # constant result
      end
      return date_count > num_count
    end

    # changes fill color of cell
    def change_fill(rgb='ffffff')
      validate_worksheet
      Color.validate_color(rgb)
      @style_index = modify_fill(@workbook, @style_index,rgb)
    end

    # Changes font name of cell
    def change_font_name(font_name='Verdana')
      validate_worksheet
      # Get copy of font object with modified name
      font = deep_copy(workbook.fonts[font_id().to_s][:font])
      font[:name][:attributes][:val] = font_name.to_s
      # Update font and xf array
      change_font(font)
    end

    # Changes font size of cell
    def change_font_size(font_size=10)
      validate_worksheet
      if font_size.is_a?(Integer) || font_size.is_a?(Float)
        # Get copy of font object with modified size
        font = deep_copy(workbook.fonts[font_id().to_s][:font])
        font[:sz][:attributes][:val] = font_size
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
      font = deep_copy(workbook.fonts[font_id().to_s][:font])
      font = modify_font_color(font, font_color.to_s)
      # Update font and xf array
      change_font(font)
    end

    # Changes font italics settings of cell
    def change_font_italics(italicized=false)
      validate_worksheet
      # Get copy of font object with modified italics settings
      font = deep_copy(workbook.fonts[font_id().to_s][:font])
      font = modify_font_italics(font, italicized)
      # Update font and xf array
      change_font(font)
    end

    # Changes font bold settings of cell
    def change_font_bold(bolded=false)
      validate_worksheet
      # Get copy of font object with modified bold settings
      font = deep_copy(workbook.fonts[font_id().to_s][:font])
      font = modify_font_bold(font, bolded)
      # Update font and xf array
      change_font(font)
    end

    # Changes font underline settings of cell
    def change_font_underline(underlined=false)
      validate_worksheet
      # Get copy of font object with modified underline settings
      font = deep_copy(workbook.fonts[font_id().to_s][:font])
      font = modify_font_underline(font, underlined)
      # Update font and xf array
      change_font(font)
    end

    # Changes font strikethrough settings of cell
    def change_font_strikethrough(struckthrough=false)
      validate_worksheet
      # Get copy of font object with modified strikethrough settings
      font = deep_copy(workbook.fonts[font_id().to_s][:font])
      font = modify_font_strikethrough(font, struckthrough)
      # Update font and xf array
      change_font(font)
    end

    # Helper method to update the font array and xf array
    def change_font(font)
      # Modify font array and retrieve new font id
      font_id = modify_font(@workbook, font, font_id())
      # Get copy of xf object with modified font id
      xf = deep_copy(xf_id())
      xf[:fontId] = Integer(font_id.to_i)
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
