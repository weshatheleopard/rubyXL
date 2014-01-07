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
      render_xml do |xml|

        i = 1

        xml << (xml.create_element('Relationships',
                :xmlns => 'http://schemas.openxmlformats.org/package/2006/relationships') { |root|

          @workbook.worksheets.each { |sheet|
            root << xml.create_element('Relationship',
              { :Id => "rId#{i}",
                :Type => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
                :Target => "worksheets/sheet#{i}.xml" })
            i += 1
          }

          unless @workbook.external_links.nil?
            1.upto(@workbook.external_links.size - 1) { |link|

              root << xml.create_element('Relationship',
                { :Id => "rId#{i}",
                  :Type => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink',
                  :Target => "externalLinks/externalLink#{link}.xml" })
              i += 1
            }
          end

          root << xml.create_element('Relationship',
            { :Id => "rId#{i}",
              :Type => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
              :Target => 'theme/theme1.xml' })
          i += 1

          root << xml.create_element('Relationship',
            { :Id => "rId#{i}",
              :Type => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
              :Target => 'styles.xml' })
          i += 1

          unless @workbook.shared_strings.empty?
            root << xml.create_element('Relationship',
              { :Id => "rId#{i}",
                :Type => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
                :Target => 'sharedStrings.xml' })
            i += 1
          end
        })
      end
    end

  end
end
end
