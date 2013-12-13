require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
 class SharedStringsWriter < GenericWriter

    def filepath
      File.join('xl', 'sharedStrings.xml')
    end

    def write()
      # Excel doesn't care much about the contents of sharedStrings.xml -- it will fill it in, but the file has to exist and have a root node.
      if @workbook.shared_strings_XML
        contents = @workbook.shared_strings_XML
      else
        contents = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+"\n"+'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="0" uniqueCount="0"></sst>'
      end
      contents
    end

  end
end
end
