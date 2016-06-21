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
