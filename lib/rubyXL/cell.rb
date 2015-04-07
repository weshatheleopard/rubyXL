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
    def change_fill(rgb = 'ffffff')
      validate_worksheet
      Color.validate_color(rgb)
      self.style_index = workbook.modify_fill(self.style_index, rgb)
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

    def change_contents(data, formula_expression = nil)
      validate_worksheet

      if formula_expression then
        self.datatype = nil
        self.formula = RubyXL::Formula.new(:expression => formula_expression)
      else
        self.datatype = case data
                        when Date, Numeric then nil
                        else RubyXL::DataType::RAW_STRING
                        end
      end

      data = workbook.date_to_num(data) if data.is_a?(Date)

      self.raw_value = data
    end

    def inspect
      str = "#<#{self.class}(#{row},#{column}): #{raw_value.inspect}"
      str += " =#{self.formula.expression}" if self.formula
      str += ", datatype = #{self.datatype}, style_index = #{self.style_index}>"
      return str
    end

    # Performs correct modification based on what type of change_type is specified
    def font_switch(change_type, arg)
      case change_type
        when Worksheet::NAME          then change_font_name(arg)
        when Worksheet::SIZE          then change_font_size(arg)
        when Worksheet::COLOR         then change_font_color(arg)
        when Worksheet::ITALICS       then change_font_italics(arg)
        when Worksheet::BOLD          then change_font_bold(arg)
        when Worksheet::UNDERLINE     then change_font_underline(arg)
        when Worksheet::STRIKETHROUGH then change_font_strikethrough(arg)
        else raise 'Invalid change_type'
      end
    end

    private

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
      raise "Cell #{self} is not in worksheet #{worksheet}"
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
