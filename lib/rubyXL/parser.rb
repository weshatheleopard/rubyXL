module RubyXL
  class Parser
    def self.parse(xl_file_path)
      root = RubyXL::WorkbookRoot.parse_file(xl_file_path)
      root.workbook
    end
  end
end
