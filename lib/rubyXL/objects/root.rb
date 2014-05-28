require 'rubyXL/objects/relationships'

module RubyXL

  class WorkbookRoot
    # Dummy object: no contents, only relationships
    attr_accessor :thumbnail, :core_properties, :document_properties, :custom_properties, :workbook
    attr_accessor :content_types, :rels_hash

    include RubyXL::RelationshipSupport

    def related_objects
      [ content_types, thumbnail, core_properties, document_properties, workbook ]
    end

    def load_relationships(dir_path)

      self.relationship_container = RubyXL::RootRelationships.load_relationship_file(dir_path, '')

      return if relationship_container.nil?

      relationship_container.load_related_files(dir_path, '').each_pair { |rid, rf|
        case rf
        when RubyXL::ThumbnailFile          then self.thumbnail = rf
        when RubyXL::CorePropertiesFile     then self.core_properties = rf
        when RubyXL::DocumentPropertiesFile then self.document_properties = rf
        when RubyXL::CustomPropertiesFile   then self.custom_properties = rf
        when RubyXL::Workbook               then self.workbook = rf
        else store_relationship(rf, :unknown)
        end
      }

    end

    def self.default
      obj = self.new
      obj.document_properties    = RubyXL::DocumentPropertiesFile.new
      obj.core_properties        = RubyXL::CorePropertiesFile.new
      obj.relationship_container = RubyXL::RootRelationships.new
      obj.content_types          = RubyXL::ContentTypes.new
      obj
    end

  end

end
