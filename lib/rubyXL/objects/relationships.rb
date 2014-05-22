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
    set_namespaces('http://schemas.openxmlformats.org/package/2006/relationships' => '')

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

    def find_by_target(target)
      relationships.find { |r| r.target == target }
    end
  end


  class WorkbookRelationships < OOXMLRelationshipsFile

    attr_accessor :workbook

    def before_write_xml
      self.relationships = []

      @workbook.worksheets.each_with_index { |sheet, i|
        relationships << document_relationship(sheet.xlsx_path.gsub(/\Axl\//, ''), sheet.rel_type)
      }

      @workbook.external_links.each_key { |k| 
        relationships << document_relationship("externalLinks/#{k}", 'externalLink')
      }

      relationships << document_relationship('theme/theme1.xml', 'theme') if @workbook.theme
      relationships << document_relationship('styles.xml', 'styles') if @workbook.stylesheet

      if @workbook.shared_strings_container && !@workbook.shared_strings_container.strings.empty? then
        relationships << document_relationship('sharedStrings.xml', 'sharedStrings') 
      end

      if @workbook.calculation_chain && !@workbook.calculation_chain.cells.empty? then
        relationships << document_relationship('calcChain.xml', 'calcChain') 
      end

      true
    end

    def self.xlsx_path
      File.join('xl', '_rels', 'workbook.xml.rels')
    end

  end

  class RootRelationships < OOXMLRelationshipsFile

    attr_accessor :workbook

    def before_write_xml
      self.relationships = []

      relationships << document_relationship('xl/workbook.xml',   'officeDocument')
      relationships << metadata_relationship('docProps/thumbnail.jpeg', 'thumbnail') unless @workbook.thumbnail.empty?
      relationships << metadata_relationship('docProps/core.xml', 'core-properties') if @workbook.core_properties 
      relationships << document_relationship('docProps/app.xml',  'extended-properties') if @workbook.document_properties

      true
    end

    def self.xlsx_path
      File.join('_rels', '.rels')
    end
  end

end
