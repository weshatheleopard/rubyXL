require 'rubygems'
require 'nokogiri'

module RubyXL
  module Writer
    class ContentTypesWriter < GenericWriter

      def filepath
        '[Content_Types].xml'
      end

      def write()

        render_xml do |xml|
          xml << (xml.create_element('Types',
                    :xmlns => 'http://schemas.openxmlformats.org/package/2006/content-types') { |root|

            unless @workbook.printer_settings.empty?
              root << xml.create_element('Default', {
                        :Extension => 'bin',
                        :ContentType => 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings' })
            end
            root << xml.create_element('Default', {
                      :Extension => 'rels',
                      :ContentType => 'application/vnd.openxmlformats-package.relationships+xml' })

            root << xml.create_element('Default', {
                      :Extension => 'xml',
                      :ContentType => 'application/xml' })

            root << xml.create_element('Default', {
                      :Extension => 'jpeg',
                      :ContentType => 'image/jpeg' })

#           if @workbook.macros.nil? && @workbook.drawings.empty?
            root << xml.create_element('Override', {
                      :PartName    => '/xl/workbook.xml',
                      :ContentType => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' })
#           else
#             xml.Override('PartName'=>'/xl/workbook.xml',
#               'ContentType'=>"application/vnd.ms-excel.sheet.macroEnabled.main+xml")
#           end

            @workbook.worksheets.each_with_index { |sheet,i|
              root << xml.create_element('Override', {
                        :PartName    => "/xl/worksheets/sheet#{i+1}.xml",
                        :ContentType => 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml' })
            }

            root << xml.create_element('Override', {
                      :PartName    => '/xl/theme/theme1.xml',
                      :ContentType => 'application/vnd.openxmlformats-officedocument.theme+xml' })

            root << xml.create_element('Override', {
                      :PartName    => '/xl/styles.xml',
                      :ContentType => 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml' })

            unless @workbook.shared_strings.empty?
              root << xml.create_element('Override', {
                        :PartName    => '/xl/sharedStrings.xml',
                        :ContentType => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml' })
            end

            @workbook.drawings.each_pair { |k, v|
              root << xml.create_element('Override', {
                        :PartName    => "/#{@workbook.drawings.local_dir_path}/#{k}",
                        :ContentType => 'application/vnd.openxmlformats-officedocument.drawing+xml' })
#            xml.Default('Extension'=>'vml',
#              'ContentType'=>'application/vnd.openxmlformats-officedocument.vmlDrawing')
            }

            @workbook.charts.each_pair { |k, v|
              case k
              when /^chart\d*.xml$/ then
                root << xml.create_element('Override', {
                          :PartName    => "/#{@workbook.charts.local_dir_path}/#{k}",
                          :ContentType => 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml' })
              when /^style\d*.xml$/ then
                root << xml.create_element('Override', {
                          :PartName    => "/#{@workbook.charts.local_dir_path}/#{k}",
                          :ContentType => 'application/vnd.ms-office.chartstyle+xml' })
              when /^colors\d*.xml$/ then
                root << xml.create_element('Override', {
                          :PartName    => "/#{@workbook.charts.local_dir_path}/#{k}",
                          :ContentType => 'application/vnd.ms-office.chartcolorstyle+xml' })
              end
            }

            root << xml.create_element('Override', {
                      :PartName    => '/docProps/core.xml',
                      :ContentType => 'application/vnd.openxmlformats-package.core-properties+xml' })

            root << xml.create_element('Override', {
                      :PartName    => '/docProps/app.xml',
                      :ContentType => 'application/vnd.openxmlformats-officedocument.extended-properties+xml' })

            unless @workbook.external_links.nil?
              1.upto(@workbook.external_links.size-1) do |i|
                root << xml.create_element('Override', {
                          :PartName    => "/xl/externalLinks/externalLink#{i}.xml",
                          :ContentType => 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml' })
              end
            end

            unless @workbook.macros.nil?
              root << xml.create_element('Override', {
                        :PartName    => '/xl/vbaProject.bin',
                        :ContentType => 'application/vnd.ms-office.vbaProject' })
            end
          })
        end
      end

    end
  end
end
