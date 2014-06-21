require 'zip'
require 'rubyXL/objects/relationships'

module RubyXL

  class WorkbookRoot
    @@debug = nil

    attr_accessor :filepath
    attr_accessor :thumbnail, :core_properties, :document_properties, :custom_properties, :workbook
    attr_accessor :content_types, :rels_hash

    REL_CLASS    = RubyXL::RootRelationships

    include RubyXL::RelationshipSupport

    def related_objects
      [ content_types, thumbnail, core_properties, document_properties, custom_properties, workbook ]
    end

#    define_relationship(RubyXL::ThumbnailFile,          :thumbnail)
#    define_relationship(RubyXL::CorePropertiesFile,     :core_properties)
#    define_relationship(RubyXL::DocumentPropertiesFile, :document_properties)
#    define_relationship(RubyXL::CustomPropertiesFile,   :custom_properties)
#    define_relationship(RubyXL::Workbook,               :workbook)

    def attach_relationship(rid, rf)
      case rf
      when RubyXL::ThumbnailFile          then self.thumbnail = rf
      when RubyXL::CorePropertiesFile     then self.core_properties = rf
      when RubyXL::DocumentPropertiesFile then self.document_properties = rf
      when RubyXL::CustomPropertiesFile   then self.custom_properties = rf
      when RubyXL::Workbook               then self.workbook = rf
      else store_relationship(rf, :unknown)
      end
    end

    def self.default
      obj = self.new
      obj.document_properties    = RubyXL::DocumentPropertiesFile.new
      obj.core_properties        = RubyXL::CorePropertiesFile.new
      obj.relationship_container = RubyXL::RootRelationships.new
      obj.content_types          = RubyXL::ContentTypes.new
      obj
    end

    def stream
      stream = Zip::OutputStream.write_buffer { |zipstream|
        self.rels_hash = {}
        self.relationship_container.owner = self
        self.content_types.overrides = []
        self.content_types.owner = self
        collect_related_objects.compact.each { |obj|
          puts "<-- DEBUG: adding relationship to #{obj.class}" if @@debug
          self.rels_hash[obj.class] ||= []
          self.rels_hash[obj.class] << obj
        }

        self.rels_hash.keys.sort_by{ |c| c.save_order }.each { |klass|
          puts "<-- DEBUG: saving related #{klass} files" if @@debug
          self.rels_hash[klass].each { |obj|
            obj.workbook = workbook if obj.respond_to?(:workbook=)
            puts "<-- DEBUG:   > #{obj.xlsx_path}" if @@debug
            self.content_types.add_override(obj)
            obj.add_to_zip(zipstream)
          }
        }
      }
      stream.rewind
      stream
    end

    def self.parse_file(xl_file_path, opts)
      begin
        ::Zip::File.open(xl_file_path) { |zip_file|
          root = self.new
          root.filepath = xl_file_path
          root.content_types = RubyXL::ContentTypes.parse_file(zip_file)
          root.load_relationships(zip_file)
          root.workbook.root = root
          root
        }
      rescue ::Zip::Error => e
        raise e, "XLSX file format error: #{e}", e.backtrace
      end
    end

  end

end
