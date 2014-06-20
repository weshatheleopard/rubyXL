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

    def add_to_zip(zip_stream)
      return if @data.nil?
      path = self.xlsx_path
      path = path.relative_path_from(Pathname.new("/")) if path.absolute? # Zip doesn't like absolute paths.
      zip_stream.put_next_entry(path)
      zip_stream.write(@data)
    end
  end

  class PrinterSettingsFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
  end

  class DrawingFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
    REL_CLASS    = RubyXL::DrawingRelationshipsFile

    include RubyXL::RelationshipSupport

    def attach_relationship(rid, rf)
      case rf
      when RubyXL::ChartFile       then store_relationship(rf) # TODO
      when RubyXL::BinaryImageFile then store_relationship(rf) # TODO
      else store_relationship(rf, :unknown)
      end
    end

  end

  class ChartFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
    REL_CLASS    = RubyXL::ChartRelationshipsFile

    include RubyXL::RelationshipSupport

    def attach_relationship(rid, rf)
      case rf
      when RubyXL::ChartColorsFile     then store_relationship(rf) # TODO
      when RubyXL::ChartStyleFile      then store_relationship(rf) # TODO
      when RubyXL::ChartUserShapesFile then store_relationship(rf) # TODO
      else store_relationship(rf, :unknown)
      end
    end

  end

  class VMLDrawingFile < GenericStorageObject
#    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'

  end

  class ChartColorsFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.ms-office.chartcolorstyle+xml'
    REL_TYPE     = 'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
  end

  class ChartStyleFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.ms-office.chartstyle+xml'
    REL_TYPE     = 'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
  end

  class TableFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
  end

  class ControlPropertiesFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp'
  end

  class PivotTableFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'
  end

  class PivotCacheDefinitionFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition'
  end

  class BinaryImageFile < GenericStorageObject
#    CONTENT_TYPE = 'image/jpeg'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
  end

  class HyperlinkRelFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
  end

  class ThumbnailFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail'
  end

  class ChartUserShapesFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chartshapes+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes'
  end

  class CustomPropertiesFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.custom-properties+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties'
  end

  class MacrosFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.ms-office.vbaProject'
    REL_TYPE     = 'http://schemas.microsoft.com/office/2006/relationships/vbaProject'
  end

  class ExternalLinksFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'
  end

  class CustomXMLFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
  end

  class SlicerFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
  end

  class SlicerCacheFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
  end

end
