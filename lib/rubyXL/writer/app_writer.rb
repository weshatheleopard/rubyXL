require 'rubygems'
require 'nokogiri'

module RubyXL
  module Writer
    class AppWriter < GenericWriter

      def filepath
        File.join('docProps', 'app.xml')
      end

      def write()
        render_xml do |xml|
          xml << (xml.create_element('Properties',
                  :xmlns => 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
                  'xmlns:vt'=>'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes') { |root|

            root << xml.create_element('Application', @workbook.application) unless @workbook.application.to_s.empty?
            root << xml.create_element('DocSecurity', 0)
            root << xml.create_element('ScaleCrop', false)

            root << (xml.create_element('HeadingPairs') { |headings|
              headings << (xml.create_element('vt:vector', :baseType => 'variant', :size => 2) { |vc|
                vc << (xml.create_element('vt:variant',) { |v| 
                  v << xml.create_element('vt:lpstr', 'Worksheets')
                })

                vc << (xml.create_element('vt:variant',) { |v|
                  v << xml.create_element('vt:i4', @workbook.worksheets.size)
                })
              })
            })

            root << (xml.create_element('TitlesOfParts') { |titles|
              titles << (xml.create_element('vt:vector', :baseType => 'lpstr',
                           :size => @workbook.worksheets.size) { |v|
                @workbook.worksheets.each { |sheet|
                  v << (xml.create_element('vt:lpstr', sheet.sheet_name))
                }
              })
            })

            root << xml.create_element('Company', @workbook.company) unless @workbook.company.to_s.empty?
            root << xml.create_element('LinksUpToDate', false)
            root << xml.create_element('SharedDoc', false)
            root << xml.create_element('HyperlinksChanged', false)
            root << xml.create_element('AppVersion', @workbook.appversion) unless @workbook.appversion.to_s.empty?
          })
        end
      end

    end
  end
end
