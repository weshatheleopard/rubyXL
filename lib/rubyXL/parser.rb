module RubyXL
  class Parser
    def self.parse(xl_file_path)
      root = RubyXL::WorkbookRoot.parse_file(xl_file_path)
      root.workbook
    end

    def self.parse_buffer(buffer)
      root = RubyXL::WorkbookRoot.parse_buffer(buffer)
      root.workbook
    end
  end
end
