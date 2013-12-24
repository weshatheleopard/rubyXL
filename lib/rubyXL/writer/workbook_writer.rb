require 'rubygems'
require 'nokogiri'

module RubyXL
  module Writer
    class WorkbookWriter < GenericWriter

      def filepath
        File.join('xl', 'workbook.xml')
      end

      def write()
        new_xml = render_xml do |xml|
          xml << (xml.create_element('workbook',
                  :xmlns => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                  'xmlns:r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') { |root|
            root << xml.create_element('fileVersion', { :appName => 'xl', :lastEdited => 4,
                                                        :lowestEdited => 4, :rupBuild => 4505 })

            params = { :showInkAnnotation => 0, :autoCompressPictures => 0 }

            #TODO following line - date 1904? check if mac only
            params[:date1904] = 1 if @workbook.date1904
            root << xml.create_element('workbookPr', params)

            root << (xml.create_element('bookViews') { |views|
              views << xml.create_element('workbookView', { :xWindow => -20, :yWindow => -20,
                                            :windowWidth => 21600, :windowHeight => 13340, :tabRatio => 500 })
            })

            index = 0
            root << (xml.create_element('sheets') { |sheet_xml|
              @workbook.worksheets.each_with_index { |sheet, i|
                index = i + 1
                sheet_xml << xml.create_element('sheet', { :name => sheet.sheet_name,
                                                           :sheetId => sheet.sheet_id || index, 
                                                           'r:id'=> "rId#{index}" })
              }
            })

            unless @workbook.external_links.empty?
              root << (xml.create_element('externalReferences') { |refs|
                # This doesn't quite make a lot of sense -- why we are starting with index and not 0?
                # Need to check once I get the file with external links...
                index.upto(@workbook.external_links.size - 1) { |id|
                  refs << xml..create_element('externalReference', { 'r:id' => "rId#{id + index}" })
                }
              })
            end

#            @workbook.defined_names.each { |name| xml << name.write_xml(xml) }
  
            #TODO see if this changes with formulas
            #attributes out of order here
            root << xml.create_element('calcPr', { :calcId => 130407, :concurrentCalc => 0 } )

            root << (xml.create_element('extLst') { |list| 
              list << (xml.create_element('ext',  {
                         'xmlns:mx' => 'http://schemas.microsoft.com/office/mac/excel/2008/main',
                         :uri => 'http://schemas.microsoft.com/office/mac/excel/2008/main'}) { |ext|
                ext << xml.create_element('mx:ArchID', { :Flags => 2 })

              })
            })
          })
        end
      end

    end
  end
end
