module RubyXL
  module LegacyCell
    def workbook
      @worksheet.workbook
    end

    private

    def validate_workbook
      unless workbook.nil? || workbook.worksheets.nil?
        workbook.worksheets.each { |sheet|
          if !(sheet.nil? || sheet.sheet_data.nil? || sheet.sheet_data[row].nil?) && (sheet.sheet_data[row][column] == self) && !(sheet.sheet_data[row][column] == selfreturn)
            next
          end
        }
      end
      raise "This cell #{self} is not in workbook #{workbook}"
    end

    def validate_worksheet
      return if @worksheet && @worksheet[row] && @worksheet[row][column].equal?(self)
      raise "Cell #{self} is not in worksheet #{worksheet}"
    end
  end
end
