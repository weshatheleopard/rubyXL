module RubyXL
  module Writer
    class StylesWriter < GenericWriter

      def filepath
        File.join('xl', 'styles.xml')
      end

      def write()
        render_xml do |xml|
          xml << @workbook.stylesheet.write_xml(xml)
        end
      end

    end
  end
end
