require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class AppWriter < GenericWriter
    FILEPATH = '/docProps/app.xml'

    def write()

      contents = build_xml do |xml|
        xml.Properties('xmlns' => 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
        'xmlns:vt'=>'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes') {
          xml.Application @workbook.application
          xml.DocSecurity '0'
          xml.ScaleCrop 'false'
          xml.HeadingPairs {
            xml['vt'].vector(:size => '2', :baseType => 'variant') {
              xml['vt'].variant {
                xml['vt'].lpstr 'Worksheets'
              }
              xml['vt'].variant {
                xml['vt'].i4 @workbook.worksheets.size.to_s()
              }
            }
          }
          xml.TitlesOfParts {
            xml['vt'].vector(:size=>@workbook.worksheets.size.to_s(), :baseType=>"lpstr") {
              @workbook.worksheets.each do |sheet|
                xml['vt'].lpstr sheet.sheet_name
              end
            }
          }
          xml.Company @workbook.company
          xml.LinksUpToDate 'false'
          xml.SharedDoc 'false'
          xml.HyperlinksChanged 'false'
          xml.AppVersion @workbook.appversion
        }
      end

      if(contents =~ /xmlns:vt=\"(.*)\" xmlns=\"(.*)\"/)
        contents.sub(/xmlns:vt=\"(.*)\" xmlns=\"(.*)\"<A/,'xmlns="'+$2+'" xmlns:vt="'+$1+'"<A')
      end

      contents
    end

  end
end
end
