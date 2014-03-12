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

    def load_related_files(zipdir_path, current_dir = '')
      self.related_files = {}

      self.relationships.each { |rel|
        file_path = Pathname.new(File.join(current_dir, rel.target)).cleanpath

        klass = RubyXL::OOXMLRelationshipsFile.get_class_by_rel_type(rel.type)

        if klass.nil? then
          puts "WARNING: storage class not found for #{rel.target} (#{rel.type})"
          klass = GenericStorageObject
        end

        self.related_files[rel.id] = klass.parse_file(zipdir_path, file_path)
      }
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

      relationships << new_relationship('theme/theme1.xml', @workbook.theme.class.rel_type)
      relationships << new_relationship('styles.xml', @workbook.stylesheet.class.rel_type)

      if @workbook.shared_strings_container && !@workbook.shared_strings_container.strings.empty? then
        relationships << new_relationship('sharedStrings.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings') 
      end

      if @workbook.calculation_chain && !@workbook.calculation_chain.cells.empty? then
        relationships << new_relationship('calcChain.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain') 
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

      relationships << new_relationship('xl/workbook.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument')
      relationships << metadata_relationship('docProps/thumbnail.jpeg', 'thumbnail') unless @workbook.thumbnail.empty?
      relationships << metadata_relationship('docProps/core.xml', 'core-properties')
      relationships << new_relationship('docProps/app.xml', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties')

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

end
