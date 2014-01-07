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
      rels = []

      @workbook.worksheets.each_index { |i|
        rels << [ "worksheets/sheet#{i + 1}.xml", 'worksheet' ]
      }

      @workbook.external_links.each_key { |k| 
        rels << [ "externalLinks/#{k}", 'externalLink' ]
      }

      rels << [ 'theme/theme1.xml', 'theme' ]
      rels << [ 'styles.xml', 'styles' ]
      rels << [ 'sharedStrings.xml', 'sharedStrings' ] unless @workbook.shared_strings.empty?

      render_xml do |xml|
        xml << (xml.create_element('Relationships',
                :xmlns => 'http://schemas.openxmlformats.org/package/2006/relationships') { |root|

          rels.each_with_index { |rel, i|
            root << xml.create_element('Relationship',
              { :Id => "rId#{i + 1}",
                :Type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/#{rel.last}",
                :Target => rel.first })
          }
        })
      end

    end

  end
end
end
