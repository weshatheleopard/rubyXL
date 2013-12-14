require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class WorkbookWriter < GenericWriter

    def filepath
      File.join('xl', 'workbook.xml')
    end

    def write()
      contents = build_xml do |xml|
        xml.workbook('xmlns'=>"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        'xmlns:r'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships") {
          #attributes out of order here
          xml.fileVersion('appName'=>'xl', 'lastEdited'=>'4','lowestEdited'=>'4','rupBuild'=>'4505')
          #TODO following line - date 1904? check if mac only
          if @workbook.date1904.nil? || @workbook.date1904.to_s == ''
            xml.workbookPr('showInkAnnotation'=>'0', 'autoCompressPictures'=>'0')
          else
            xml.workbookPr('date1904'=>@workbook.date1904.to_s, 'showInkAnnotation'=>'0', 'autoCompressPictures'=>'0')
          end
          xml.bookViews {
            #attributes out of order here
            xml.workbookView('xWindow'=>'-20', 'yWindow'=>'-20',
              'windowWidth'=>'21600','windowHeight'=>'13340','tabRatio'=>'500')
          }

          index = 0
          xml.sheets do
            @workbook.worksheets.each_with_index { |sheet, i|
              xml.sheet('name'=> sheet.sheet_name,
                        'sheetId' => (i + 1).to_s,
                        'r:id'=> "rId#{i + 1}")
              index = i + 1
            }
          end

          unless @workbook.external_links.empty?
            xml.externalReferences do
              # This doesn't quite make a lot of sense -- why we are starting with index and not 0?
              # Need to check once I get the file with external links...
              index.upto(@workbook.external_links.size - 1) { |id|
                xml.externalReference('r:id' => "rId#{id + index}")
              }
            end
          end

          # nokogiri builder creates CDATA tag around content,
          # using .text creates "html safe" &lt; and &gt; in place of < and >
          # xml to hash method does not seem to function well for this particular piece of xml
          xml.cdata @workbook.defined_names.to_s

          #TODO see if this changes with formulas
          #attributes out of order here
          xml.calcPr('calcId'=>'130407', 'concurrentCalc'=>'0')
          xml.extLst {
            xml.ext('xmlns:mx'=>"http://schemas.microsoft.com/office/mac/excel/2008/main",
            'uri'=>"http://schemas.microsoft.com/office/mac/excel/2008/main") {
              xml['mx'].ArchID('Flags'=>'2')
            }
          }
        }
      end

      # See comment above about CDATA. This needs to be implemented properly.
      contents.gsub(/<!\[CDATA\[(.*)\]\]>/, '\1')

    end

  end
end
end
