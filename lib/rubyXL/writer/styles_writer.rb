module RubyXL
  module Writer
    class StylesWriter < GenericWriter

      def filepath
        File.join('xl', 'styles.xml')
      end

      def ooxml_object
        @workbook.stylesheet
      end

    end
  end
end
