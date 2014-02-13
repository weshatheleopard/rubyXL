require 'rubyXL/objects/ooxml_object'

module RubyXL

  class RID < OOXMLObject
    define_attribute(:'r:id',            :string, :required => true)
  end

  class Relationship < OOXMLObject
    define_attribute(:Id,     :string)
    define_attribute(:Type,   :string)
    define_attribute(:Target, :string)
    define_element_name 'Relationship'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_calcChain.html
  class WorkbookRelationships < OOXMLTopLevelObject
    define_child_node(RubyXL::Relationship, :collection => true, :accessor => :relationships)
    define_element_name 'Relationships'
    set_namespaces('xmlns' => 'http://schemas.openxmlformats.org/package/2006/relationships')

    attr_accessor :workbook

    def create_relationship(target, type)
      RubyXL::Relationship.new(:id => "rId#{relationships.size + 1}", 
                               :type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/#{type}",
                               :target => target)
    end

    def before_write_xml
      self.relationships = []

      @workbook.worksheets.each_with_index { |sheet, i|
        relationships << create_relationship(sheet.filepath.gsub(/^xl\//, ''), sheet.rel_type)
      }

      @workbook.external_links.each_key { |k| 
        relationships << create_relationship("externalLinks/#{k}", 'externalLink')
      }

      relationships << create_relationship('theme/theme1.xml', 'theme')
      relationships << create_relationship('styles.xml', 'styles')

      if @workbook.shared_strings_container && !@workbook.shared_strings_container.strings.empty? then
        relationships << create_relationship('sharedStrings.xml', 'sharedStrings') 
      end

      if @workbook.calculation_chain && !@workbook.calculation_chain.cells.empty? then
        relationships << create_relationship('calcChain.xml', 'calcChain') 
      end

      true
    end

    def find_by_rid(r_id)
      relationships.find { |r| r.id == r_id }
    end

    def self.filepath
      File.join('xl', '_rels', 'workbook.xml.rels')
    end
  end
end
