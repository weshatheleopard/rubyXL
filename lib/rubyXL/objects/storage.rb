module RubyXL

  class GenericStorageObject

    attr_accessor :xlsx_path, :data, :workbook, :generic_storage

    def initialize
      @workbook = nil
      @xlsx_path = nil
      @data = nil
      @generic_storage = []
    end

    def self.save_order
      0
    end

    def self.parse_file(dirpath, file_path = nil)
      full_path = File.join(dirpath, file_path || self.xlsx_path)
      return nil unless File.exist?(full_path)

      obj = self.new
      obj.xlsx_path = file_path
      obj.data = File.open(full_path, 'r').read
      obj
    end

    def add_to_zip(zipfile)
      return if @data.nil?

      path = self.xlsx_path
      path = path.relative_path_from(Pathname.new("/")) if path.absolute? # Zip doesn't like absolute paths.

      zipfile.get_output_stream(path) { |f| f << @data }
    end
  end

  class PrinterSettingsFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
    end
  end

  class DrawingFile < GenericStorageObject
    include RubyXL::RelationshipSupport

    def load_relationships(dir_path, base_file_name)

      self.relationship_container = RubyXL::DrawingRelationships.load_relationship_file(dir_path, base_file_name)

      return if relationship_container.nil?

      relationship_container.load_related_files(dir_path, base_file_name)

      related_files = relationship_container.related_files
      related_files.each_pair { |rid, rf|
        case rf
        when RubyXL::ChartFile       then store_relationship(rf) # TODO
        when RubyXL::BinaryImageFile then store_relationship(rf) # TODO
        else store_relationship(rf, :unknown)
        end
      }

    end

    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
    end

    def self.content_type
      'application/vnd.openxmlformats-officedocument.drawing+xml'
    end

  end

  class ChartFile < GenericStorageObject
    attr_accessor :relationship_container

    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
    end

    def self.content_type
      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    end

    def load_relationships(dir_path, base_file_name)

      self.relationship_container = RubyXL::ChartRelationships.load_relationship_file(dir_path, base_file_name)

      if relationship_container then
        relationship_container.load_related_files(dir_path, base_file_name)

        related_files = relationship_container.related_files
        related_files.each_pair { |rid, rf|
          case rf
          when RubyXL::ChartColorsFile     then self.generic_storage << rf # TODO
          when RubyXL::ChartStyleFile      then self.generic_storage << rf # TODO
          when RubyXL::ChartUserShapesFile then self.generic_storage << rf # TODO
          else
            self.generic_storage << rf
puts "-! DEBUG: #{self.class}: unattached: #{rf.class}"
          end
        }
      end
    end

    include RubyXL::RelationshipSupport

  end

  class VMLDrawingFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
    end

#    def self.content_type
#      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
#    end
  end

  class ChartColorsFile < GenericStorageObject
    def self.rel_type
      'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
    end

    def self.content_type
      'application/vnd.ms-office.chartcolorstyle+xml'
    end
  end

  class ChartStyleFile < GenericStorageObject
    def self.rel_type
      'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
    end

    def self.content_type
      'application/vnd.ms-office.chartstyle+xml'
    end
  end

  class TableFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
    end
  end

  class ControlPropertiesFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp'
    end
  end

  class PivotTableFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'
    end
  end

  class PivotCacheDefinitionFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition'
    end
  end

  class BinaryImageFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
    end

#    def self.content_type
#      'image/jpeg'
#    end
  end

  class HyperlinkRelFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
    end
  end

  class ThumbnailFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail'
    end

    # default path = 'docProps/thumbnail.jpeg'
  end

  class ChartUserShapesFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes'
    end

    def self.content_type
      'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml'
    end
  end

  class CustomPropertiesFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
    end

    def self.content_type
      'application/vnd.openxmlformats-officedocument.custom-properties+xml'
    end
  end

  class MacrosFile < GenericStorageObject
#    def self.rel_type
#      'http://schemas...'
#    end

    def self.content_type
      'application/vnd.ms-office.vbaProject'
    end
  end

  class ExternalLinksFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
    end

    def self.content_type
      'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'
    end
  end

  class CustomXMLFile < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
    end
  end

end
