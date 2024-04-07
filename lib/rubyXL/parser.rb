module RubyXL
  class Parser
    # Parse <tt>.xslx</tt> file by reading it from local disk.
    def self.parse(src_file_path)
      begin
        ::Zip::File.open(src_file_path) { |zip_file|
          root = RubyXL::WorkbookRoot.parse_zip_file(zip_file)
          root.source_file_path = src_file_path
          root.workbook
        }
      rescue ::Zip::Error => e
        raise e, "XLSX file format error: #{e}", e.backtrace
      end
    end

    # Parse <tt>.xslx</tt> file contained in a stream (useful for receiving over HTTP).
    def self.parse_buffer(buffer)
      begin
        zip_file = ::Zip::File.open_buffer(buffer)
        root = RubyXL::WorkbookRoot.parse_zip_file(zip_file)
        root.workbook
      rescue ::Zip::Error => e
        raise e, "XLSX file format error: #{e}", e.backtrace
      end
    end
  end
end
