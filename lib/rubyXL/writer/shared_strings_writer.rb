module RubyXL
  module Writer
   class SharedStringsWriter < GenericWriter

      def filepath
        File.join('xl', 'sharedStrings.xml')
      end

      def ooxml_object
        @workbook.shared_strings_container
      end

    end
  end
end
