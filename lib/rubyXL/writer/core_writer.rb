require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class CoreWriter < GenericWriter

    def filepath
      File.join('docProps', 'core.xml')
    end

    def write()
      render_xml do |xml|
        xml << (xml.create_element('cp:coreProperties', 
                   'xmlns:cp' => 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                   'xmlns:dc' => 'http://purl.org/dc/elements/1.1/',
                   'xmlns:dcterms' => 'http://purl.org/dc/terms/',
                   'xmlns:dcmitype' => 'http://purl.org/dc/dcmitype/',
                   'xmlns:xsi' => 'http://www.w3.org/2001/XMLSchema-instance') { |root|

          root << xml.create_element('dc:creator',        @workbook.creator)
          root << xml.create_element('cp:lastModifiedBy', @workbook.modifier)

          unless @workbook.created_at.to_s.empty?
            root << xml.create_element('dcterms:created',  { 'xsi:type' => 'dcterms:W3CDTF' }, @workbook.created_at)
          end

          unless @workbook.modified_at.to_s.empty?
            root << xml.create_element('dcterms:modified', { 'xsi:type' => 'dcterms:W3CDTF' }, @workbook.modified_at)
          end
        })
      end
    end

  end
end
end
