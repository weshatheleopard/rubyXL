require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer

  class WorkbookRelsWriter < GenericWriter

    def filepath
      File.join('xl', '_rels', 'workbook.xml.rels')
    end

    #all attributes out of order
    def write()
      build_xml do |xml|
        xml.Relationships('xmlns'=>'http://schemas.openxmlformats.org/package/2006/relationships') {
          i = 1
          @workbook.worksheets.each do |sheet|
            xml.Relationship('Id'=>'rId'+i.to_s,
              'Type'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
              'Target'=>'worksheets/sheet'+i.to_s+'.xml')
            i += 1
          end
          unless @workbook.external_links.nil?
            1.upto(@workbook.external_links.size-1) do |link|
              xml.Relationship('Id'=>'rId'+i.to_s,
                'Type'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
                'Target'=>"externalLinks/externalLink#{link}.xml"
              )
              i+=1
            end
          end
          xml.Relationship('Id'=>'rId'+i.to_s,
            'Type'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            'Target'=>'theme/theme1.xml')
          i += 1
          xml.Relationship('Id'=>'rId'+i.to_s,
            'Type'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            'Target'=>'styles.xml')

          unless @workbook.shared_strings.empty?
            i+=1
            xml.Relationship('Id'=>'rId'+i.to_s,
            'Type'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
              'Target'=>'sharedStrings.xml')
          end
        }
      end
    end

  end
end
end
