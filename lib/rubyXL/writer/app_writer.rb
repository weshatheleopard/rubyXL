module RubyXL
  module Writer
    class AppWriter < GenericWriter

      def filepath
        File.join('docProps', 'app.xml')
      end

      def ooxml_object
        @workbook.document_properties.workbook = @workbook
        @workbook.document_properties
      end

    end
  end
end
