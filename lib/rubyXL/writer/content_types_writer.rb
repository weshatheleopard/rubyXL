require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class ContentTypesWriter < GenericWriter
    FILEPATH = '/[Content_Types].xml'

    def write()
      build_xml do |xml|
        xml.Types('xmlns'=>"http://schemas.openxmlformats.org/package/2006/content-types") {
          xml.Default('Extension'=>'xml', 'ContentType'=>'application/xml')
          unless @workbook.shared_strings.empty?
            xml.Override('PartName'=>'/xl/sharedStrings.xml',
              'ContentType'=>"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml")
          end
          if @workbook.macros.nil? && @workbook.drawings.nil?
            xml.Override('PartName'=>'/xl/workbook.xml',
              'ContentType'=>"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
          else
            xml.Override('PartName'=>'/xl/workbook.xml',
              'ContentType'=>"application/vnd.ms-excel.sheet.macroEnabled.main+xml")
          end
          xml.Override('PartName'=>"/xl/styles.xml",
            'ContentType'=>"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
          xml.Default('Extension'=>'rels','ContentType'=>'application/vnd.openxmlformats-package.relationships+xml')
          unless @workbook.external_links.nil?
            1.upto(@workbook.external_links.size-1) do |i|
              xml.Override('PartName'=>"/xl/externalLinks/externalLink#{i}.xml",
              'ContentType'=>"application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml")
            end
          end
          unless @workbook.macros.nil?
            xml.Override('PartName'=>'/xl/vbaProject.bin',
              'ContentType'=>'application/vnd.ms-office.vbaProject')
          end
          unless @workbook.printer_settings.nil?
            xml.Default('Extension'=>'bin',
              'ContentType'=>'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings')
          end
          unless @workbook.drawings.nil?
            xml.Default('Extension'=>'vml',
              'ContentType'=>'application/vnd.openxmlformats-officedocument.vmlDrawing')
          end
          xml.Override('PartName'=>'/xl/theme/theme1.xml',
            'ContentType'=>"application/vnd.openxmlformats-officedocument.theme+xml")
          @workbook.worksheets.each_with_index do |sheet,i|
            xml.Override('PartName'=>"/xl/worksheets/sheet#{i+1}.xml",
              'ContentType'=>"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
          end
          xml.Override('PartName'=>'/docProps/core.xml',
            'ContentType'=>"application/vnd.openxmlformats-package.core-properties+xml")
          xml.Override('PartName'=>'/docProps/app.xml',
            'ContentType'=>"application/vnd.openxmlformats-officedocument.extended-properties+xml")
        }
      end

    end

  end
end
end
