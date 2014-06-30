module RubyXL

  # http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.cellvalues(v=office.14).aspx
  module DataType
    SHARED_STRING = 's'
    RAW_STRING    = 'str'
    INLINE_STRING = 'inlineStr'
    ERROR         = 'e'
    BOOLEAN       = 'b'
    NUMBER        = 'n'
    DATE          = 'd'  # Only available in Office2010.
  end

  module LegacyCell
    attr_accessor :formula, :worksheet

    def workbook
      @worksheet.workbook
    end

    # changes fill color of cell
    def change_fill(rgb='ffffff')
      validate_worksheet
      Color.validate_color(rgb)
      self.style_index = workbook.modify_fill(self.style_index,rgb)
    end

    # Changes font name of cell
    def change_font_name(new_font_name = 'Verdana')
      validate_worksheet

      font = get_cell_font.dup
      font.set_name(new_font_name)
      update_font_references(font)
    end

    # Changes font size of cell
    def change_font_size(font_size = 10)
      validate_worksheet
      raise 'Argument must be a number' unless font_size.is_a?(Integer) || font_size.is_a?(Float)

      font = get_cell_font.dup
      font.set_size(font_size)
      update_font_references(font)
    end

    # Changes font color of cell
    def change_font_color(font_color = '000000')
      validate_worksheet
      Color.validate_color(font_color)

      font = get_cell_font.dup
      font.set_rgb_color(font_color)
      update_font_references(font)
    end

    # Changes font italics settings of cell
    def change_font_italics(italicized = false)
      validate_worksheet

      font = get_cell_font.dup
      font.set_italic(italicized)
      update_font_references(font)
    end

    # Changes font bold settings of cell
    def change_font_bold(bolded = false)
      validate_worksheet

      font = get_cell_font.dup
      font.set_bold(bolded)
      update_font_references(font)
    end

    # Changes font underline settings of cell
    def change_font_underline(underlined = false)
      validate_worksheet

      font = get_cell_font.dup
      font.set_underline(underlined)
      update_font_references(font)
    end

    def change_font_strikethrough(struckthrough = false)
      validate_worksheet

      font = get_cell_font.dup
      font.set_strikethrough(struckthrough)
      update_font_references(font)
    end

    # Helper method to update the font array and xf array
    def update_font_references(modified_font)
      xf = workbook.register_new_font(modified_font, get_cell_xf)
      self.style_index = workbook.register_new_xf(xf, self.style_index)
    end
    private :update_font_references

    # changes horizontal alignment of cell
    def change_horizontal_alignment(alignment = 'center')
      validate_worksheet
      self.style_index = workbook.modify_alignment(self.style_index, true, alignment)
    end

    # changes vertical alignment of cell
    def change_vertical_alignment(alignment = 'center')
      validate_worksheet
      self.style_index = workbook.modify_alignment(self.style_index, false, alignment)
    end

    # changes wrap of cell
    def change_text_wrap(wrap = false)
      validate_worksheet
      self.style_index = workbook.modify_text_wrap(self.style_index, wrap)
    end

    def change_border(direction, weight)
      validate_worksheet
      self.style_index = workbook.modify_border(self.style_index, direction, weight)
    end

    def change_border_top(weight = 'thin')
      change_border(:top, weight)
    end

    def change_border_left(weight = 'thin')
      change_border(:left, weight)
    end

    def change_border_right(weight = 'thin')
      change_border(:right, weight)
    end

    def change_border_bottom(weight = 'thin')
      change_border(:bottom, weight)
    end

    def change_border_diagonal(weight = 'thin')
      change_border(:diagonal, weight)
    end

    def change_contents(data, formula_expression = nil)
      validate_worksheet

      if formula_expression then
        self.datatype = nil
        self.formula = RubyXL::Formula.new(:expression => formula_expression)
      else
        self.datatype = case data
                        when Date, Integer, Float then nil
                        else RubyXL::DataType::RAW_STRING
                        end
      end

      data = workbook.date_to_num(data) if data.is_a?(Date)

      self.raw_value = data
    end

    # returns if font is italicized
    def is_italicized()
      validate_worksheet
      get_cell_font.is_italic
    end

    # returns if font is bolded
    def is_bolded()
      validate_worksheet
      get_cell_font.is_bold
    end

    def is_underlined()
      validate_worksheet
      get_cell_font.is_underlined
    end

    def is_struckthrough()
      validate_worksheet
      get_cell_font.is_strikethrough
    end

    def font_name()
      validate_worksheet
      get_cell_font.get_name
    end

    def font_size()
      validate_worksheet
      get_cell_font.get_size
    end

    def font_color()
      validate_worksheet
      get_cell_font.get_rgb_color || '000000'
    end

    # returns cell's fill color
    def fill_color()
      validate_worksheet
      return workbook.get_fill_color(get_cell_xf)
    end

    # returns cell's horizontal alignment
    def horizontal_alignment()
      validate_worksheet
      xf_obj = get_cell_xf
      return nil if xf_obj.alignment.nil?
      xf_obj.alignment.horizontal
    end

    # returns cell's vertical alignment
    def vertical_alignment()
      validate_worksheet
      xf_obj = get_cell_xf
      return nil if xf_obj.alignment.nil?
      xf_obj.alignment.vertical
    end

    # returns cell's wrap
    def text_wrap()
      validate_worksheet
      xf_obj = get_cell_xf
      return nil if xf_obj.alignment.nil?
      xf_obj.alignment.wrap_text
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
      str = "#<#{self.class}(#{row},#{column}): #{raw_value.inspect}" 
      str += " =#{self.formula.expression}" if self.formula
      str += ", datatype = #{self.datatype}, style_index = #{self.style_index}>"
      return str
    end

    private

    def get_border(direction)
      validate_worksheet
      get_cell_border.get_edge_style(direction)
    end

    def validate_workbook()
      unless workbook.nil? || workbook.worksheets.nil?
        workbook.worksheets.each { |sheet|
          unless sheet.nil? || sheet.sheet_data.nil? || sheet.sheet_data[row].nil?
            if sheet.sheet_data[row][column] == self
              return
            end
          end
        }
      end
      raise "This cell #{self} is not in workbook #{workbook}"
    end

    def validate_worksheet()
      return if @worksheet && @worksheet[row] && @worksheet[row][column] == self
      raise "This cell #{self} is not in worksheet #{worksheet}"
    end

    def get_cell_xf
      workbook.cell_xfs[self.style_index || 0]
    end

    def get_cell_font
      workbook.fonts[get_cell_xf.font_id]
    end

    def get_cell_border
      workbook.borders[get_cell_xf.border_id]
    end

  end
end
