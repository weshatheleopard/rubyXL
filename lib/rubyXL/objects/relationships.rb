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

  class OOXMLRelationshipsFile < OOXMLTopLevelObject
    define_child_node(RubyXL::Relationship, :collection => true, :accessor => :relationships)
    define_element_name 'Relationships'
    set_namespaces('xmlns' => 'http://schemas.openxmlformats.org/package/2006/relationships')

    def document_relationship(target, type)
      RubyXL::Relationship.new(:id => "rId#{relationships.size + 1}", 
                               :type => "http://schemas.openxmlformats.org/officeDocument/2006/relationships/#{type}",
                               :target => target)
    end
    protected :document_relationship

    def metadata_relationship(target, type)
      RubyXL::Relationship.new(:id => "rId#{relationships.size + 1}", 
                               :type => "http://schemas.openxmlformats.org/package/2006/relationships/metadata/#{type}",
                               :target => target)
    end
    protected :metadata_relationship

    def find_by_rid(r_id)
      relationships.find { |r| r.id == r_id }
    end
  end


  class WorkbookRelationships < OOXMLRelationshipsFile

    attr_accessor :workbook

    def before_write_xml
      self.relationships = []

      @workbook.worksheets.each_with_index { |sheet, i|
        relationships << document_relationship(sheet.filepath.gsub(/^xl\//, ''), sheet.rel_type)
      }

      @workbook.external_links.each_key { |k| 
        relationships << document_relationship("externalLinks/#{k}", 'externalLink')
      }

      relationships << document_relationship('theme/theme1.xml', 'theme')
      relationships << document_relationship('styles.xml', 'styles')

      if @workbook.shared_strings_container && !@workbook.shared_strings_container.strings.empty? then
        relationships << document_relationship('sharedStrings.xml', 'sharedStrings') 
      end

      if @workbook.calculation_chain && !@workbook.calculation_chain.cells.empty? then
        relationships << document_relationship('calcChain.xml', 'calcChain') 
      end

      true
    end

    def self.filepath
      File.join('xl', '_rels', 'workbook.xml.rels')
    end

  end

  class RootRelationships < OOXMLRelationshipsFile

    attr_accessor :workbook

    def before_write_xml
      self.relationships = []

      relationships << document_relationship('xl/workbook.xml', 'officeDocument')
      relationships << metadata_relationship('docProps/core.xml', 'core-properties')
      relationships << document_relationship('docProps/app.xml', 'extended-properties')

      true
    end

    def self.filepath
      File.join('_rels', '.rels')
    end
  end

end
