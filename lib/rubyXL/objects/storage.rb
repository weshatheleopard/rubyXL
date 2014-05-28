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
      file_path ||= self.xlsx_path
      obj = self.new

      case dirpath
      when String then
        full_path = File.join(dirpath, file_path)
        return nil unless File.exist?(full_path)
        obj.data = File.open(full_path, 'r') { |f| f.read }
      when Zip::File then
        entry = dirpath.find_entry(file_path)
        return nil if entry.nil?
        obj.data = entry.get_input_stream { |f| f.read }
      end

      obj.xlsx_path = file_path
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

    def relationship_file_class
      RubyXL::DrawingRelationshipsFile
    end

    def attach_relationship(rid, rf)
      case rf
      when RubyXL::ChartFile       then store_relationship(rf) # TODO
      when RubyXL::BinaryImageFile then store_relationship(rf) # TODO
      else store_relationship(rf, :unknown)
      end
    end

    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
    end

    def self.content_type
      'application/vnd.openxmlformats-officedocument.drawing+xml'
    end

  end

  class ChartFile < GenericStorageObject
    include RubyXL::RelationshipSupport

    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
    end

    def self.content_type
      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    end

    def relationship_file_class
      RubyXL::ChartRelationshipsFile
    end

    def attach_relationship(rid, rf)
      case rf
      when RubyXL::ChartColorsFile     then self.generic_storage << rf # TODO
      when RubyXL::ChartStyleFile      then self.generic_storage << rf # TODO
      when RubyXL::ChartUserShapesFile then self.generic_storage << rf # TODO
      else store_relationship(rf, :unknown)
      end
    end

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
    def self.rel_type
      'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
    end

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
