require 'rubyXL/objects/ooxml_object'

module RubyXL

  class RID < OOXMLObject
    define_attribute(:'r:id',            :string, :required => true)
  end

  class Relationship < OOXMLObject
    define_attribute(:Id,         :string)
    define_attribute(:Type,       :string)
    define_attribute(:Target,     :string)
    define_attribute(:TargetMode, :string)
    define_element_name 'Relationship'
  end

  class OOXMLRelationshipsFile < OOXMLTopLevelObject
    define_child_node(RubyXL::Relationship, :collection => true, :accessor => :relationships)
    define_element_name 'Relationships'
    set_namespaces('http://schemas.openxmlformats.org/package/2006/relationships' => '')

    attr_accessor :related_files, :owner

    def new_relationship(target, type)
      RubyXL::Relationship.new(:id => "rId#{relationships.size + 1}", 
                               :type => type,
                               :target => target)
    end
    protected :new_relationship

    def add_relationship(obj)
      return if obj.nil?
      relationships << RubyXL::Relationship.new(:id => "rId#{relationships.size + 1}", 
                                                :type => obj.class.rel_type,
                                                :target => obj.xlsx_path)
    end
    protected :add_relationship

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

@@debug_depth +=2

      self.relationships.each { |rel|
        next if rel.target_mode == 'External'

        file_path = Pathname.new(rel.target)

        if !file_path.absolute? then
          file_path = (Pathname.new(File.dirname(base_file_name)) + file_path).cleanpath
        end

        klass = RubyXL::OOXMLRelationshipsFile.get_class_by_rel_type(rel.type)

        if klass.nil? then
          puts "*** WARNING: storage class not found for #{rel.target} (#{rel.type})"
          klass = GenericStorageObject
        end

puts "==>DEBUG:#{'  ' * @@debug_depth}Loading #{klass} (#{rel.id}): #{file_path}"

        obj = klass.parse_file(zipdir_path, file_path)
        obj.load_relationships(zipdir_path, file_path) if obj.respond_to?(:load_relationships)
        self.related_files[rel.id] = obj
      }
@@debug_depth -=2

      related_files
    end

    def self.load_relationship_file(zipdir_path, base_file_path)
      rel_file_path = Pathname.new(File.join(File.dirname(base_file_path), '_rels', File.basename(base_file_path) + '.rels')).cleanpath

@@debug_depth = 0 unless defined?(@@debug_depth)
puts "==>DEBUG:  #{'  ' * @@debug_depth}Loading .rel file: #{rel_file_path}"

      parse_file(zipdir_path, rel_file_path)
    end

    def xlsx_path
      file_path = owner.xlsx_path
      File.join(File.dirname(file_path), '_rels', File.basename(file_path) + '.rels')
    end

  end
	
  class WorkbookRelationships < OOXMLRelationshipsFile

    attr_accessor :workbook

    def before_write_xml
      self.relationships = []

      @workbook.worksheets.each_with_index { |sheet, i|
        relationships << new_relationship(sheet.xlsx_path.gsub(/\Axl\//, ''), sheet.class.rel_type)
      }

#      @workbook.external_links.each_key { |k| 
#        relationships << new_relationship("externalLinks/#{k}", 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink')
#      }

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

  end

  class RootRelationships < OOXMLRelationshipsFile

    def before_write_xml
      self.relationships = []

      add_relationship(owner.workbook)
      add_relationship(owner.thumbnail)
      add_relationship(owner.core_properties)
      add_relationship(owner.document_properties)

      true
    end

    def xlsx_path
      File.join('_rels', '.rels')
    end
  end

  class SheetRelationshipsFile < OOXMLRelationshipsFile
    # Insert class specific stuff here once we get to implementing it
  end

  class DrawingRelationshipsFile < OOXMLRelationshipsFile
    # Insert class specific stuff here once we get to implementing it
  end

  class ChartRelationshipsFile < OOXMLRelationshipsFile
    # Insert class specific stuff here once we get to implementing it
  end

  module RelationshipSupport

    attr_accessor :generic_storage, :relationship_container
    def related_objects # Override this method
      []
    end

    def collect_related_objects
      res = related_objects.compact # Avoid tainting +related_objects+ array

      res += generic_storage if generic_storage

      if relationship_container then
        relationship_container.owner = self
        res << relationship_container
      end

      res.each { |o| res += o.collect_related_objects if o.respond_to?(:collect_related_objects) }

      res
    end

    def store_relationship(related_file, unknown = false)
      self.generic_storage ||= []
puts "-! DEBUG: #{self.class}: unattached: #{related_file.class}" if unknown
      self.generic_storage << related_file
    end

  end

end
