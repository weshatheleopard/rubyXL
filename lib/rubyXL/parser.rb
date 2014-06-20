require 'nokogiri'
require 'tmpdir'
require 'zip'

module RubyXL

  class Parser
    def self.parse(file_path, opts = {})
      self.new(opts).parse(file_path)
    end

    # +:data_only+ allows only the sheet data to be parsed, so as to speed up parsing.
    # However, using this option will result in date-formatted cells being interpreted as numbers.
    def initialize(opts = {})
      @opts = opts
    end

    def parse(xl_file_path)
      root = RubyXL::WorkbookRoot.parse_file(xl_file_path, @opts)

      wb = root.workbook

      wb.sheets.each_with_index { |sheet, i|
        sheet_obj = wb.relationship_container.related_files[sheet.r_id]

        wb.worksheets[i] = sheet_obj # Must be done first so the sheet becomes aware of its number
        sheet_obj.workbook = wb

        sheet_obj.sheet_name = sheet.name
        sheet_obj.sheet_id = sheet.sheet_id
        sheet_obj.state = sheet.state
      }

      wb
    end

  end
end
