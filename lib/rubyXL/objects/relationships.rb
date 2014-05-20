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

    attr_accessor :related_files

    def new_relationship(target, type)
      RubyXL::Relationship.new(:id => "rId#{relationships.size + 1}", 
                               :type => type,
                               :target => target)
    end
    protected :new_relationship

    def metadata_relationship(target, type)
      RubyXL::Relationship.new(:id => "rId#{relationships.size + 1}", 
                               :type => "http://schemas.openxmlformats.org/package/2006/relationships/metadata/#{type}",
                               :target => target)
    end
    protected :metadata_relationship

    def self.save_order
      0
    end

    def find_by_rid(r_id)
      relationships.find { |r| r.id == r_id }
    end

    def find_by_target(target)
      relationships.find { |r| r.target == target }
    end

    def self.get_class_by_rel_type(rel_type)
      unless defined?(@@rel_hash)
        @@rel_hash = {}
        RubyXL.constants.each { |c|
          klass = RubyXL.const_get(c)
          @@rel_hash[klass.rel_type] = klass if klass.respond_to?(:rel_type)
        }
      end

      @@rel_hash[rel_type]
    end

    def load_related_files(zipdir_path, base_file_name = '')
      self.related_files = {}

      self.relationships.each { |rel|
        file_path = Pathname.new(rel.target)

        if !file_path.absolute? then
          file_path = (Pathname.new(File.dirname(base_file_name)) + file_path).cleanpath
        end

        klass = RubyXL::OOXMLRelationshipsFile.get_class_by_rel_type(rel.type)

        if klass.nil? then
          puts "WARNING: storage class not found for #{rel.target} (#{rel.type})"
          klass = GenericStorageObject
        end

puts ">>>DEBUG: Loading related file: rid=#{rel.id} path=#{file_path} klass=#{klass}"

        obj = klass.parse_file(zipdir_path, file_path)
        obj.load_relationships(zipdir_path, file_path) if obj.respond_to?(:load_relationships)
        self.related_files[rel.id] = obj
      }
    end

    def self.load_relationship_file(zipdir_path, base_file_path)
      rel_file_path = File.join(File.dirname(base_file_path), '_rels', File.basename(base_file_path) + '.rels')

puts ">>>DEBUG: Loading .rel file: base_file=#{base_file_path} rel_file=#{rel_file_path}"

      parse_file(zipdir_path, rel_file_path)
    end

  end
	
  class WorkbookRelationships < OOXMLRelationshipsFile

    attr_accessor :workbook

    def before_write_xml
      self.relationships = []

      @workbook.worksheets.each_with_index { |sheet, i|
        relationships << new_relationship(sheet.xlsx_path.gsub(/^xl\//, ''), sheet.class.rel_type)
      }

      @workbook.external_links.each_key { |k| 
        relationships << new_relationship("externalLinks/#{k}", 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink')
      }

      relationships << new_relationship('theme/theme1.xml', @workbook.theme.class.rel_type) if @workbook.theme
      relationships << new_relationship('styles.xml', @workbook.stylesheet.class.rel_type) if @workbook.stylesheet

      if @workbook.shared_strings_container && !@workbook.shared_strings_container.strings.empty? then
        relationships << new_relationship('sharedStrings.xml', @workbook.shared_strings_container.class.rel_type)
      end

      if @workbook.calculation_chain && !@workbook.calculation_chain.cells.empty? then
        relationships << new_relationship('calcChain.xml', @workbook.calculation_chain.class.rel_type)
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

      relationships << new_relationship('xl/workbook.xml', @workbook.class.rel_type)
      relationships << metadata_relationship('docProps/thumbnail.jpeg', 'thumbnail') unless @workbook.thumbnail.empty?
      relationships << new_relationship('docProps/core.xml',@workbook.core_properties.class.rel_type) if @workbook.core_properties 
      relationships << new_relationship('docProps/app.xml', RubyXL::DocumentProperties.rel_type) if @workbook.document_properties

      true
    end

    def self.xlsx_path
      File.join('_rels', '.rels')
    end
  end

  class SheetRelationships < OOXMLRelationshipsFile

    attr_accessor :sheet

    def initialize(params = {})
      super
      self.sheet = params[:sheet]
    end

    def xlsx_path
      file_path = sheet.xlsx_path
      File.join(File.dirname(file_path), '_rels', File.basename(file_path) + '.rels')
    end

  end

  class DrawingRelationships < OOXMLRelationshipsFile
    attr_accessor :owner

#    def initialize(params = {})
#      super
#      self.owner = params[:owner]
#    end

    def xlsx_path
      file_path = owner.xlsx_path
      File.join(File.dirname(file_path), '_rels', File.basename(file_path) + '.rels')
    end

  end


  class ChartRelationships < OOXMLRelationshipsFile
    attr_accessor :owner

#    def initialize(params = {})
#      super
#      self.owner = params[:owner]
#    end

    def xlsx_path
      file_path = owner.xlsx_path
      File.join(File.dirname(file_path), '_rels', File.basename(file_path) + '.rels')
    end

  end


  module RelationshipSupport
    def collect_related_objects
      res = [] + related_objects
      related_objects.each { |o| res += o.collect_related_objects if o.respond_to?(:collect_related_objects) }
      res
    end
  end

end
