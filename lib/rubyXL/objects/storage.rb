module RubyXL

  class GenericStorageObject
    SAVE_ORDER = 0

    attr_accessor :xlsx_path, :data, :generic_storage

    def initialize(file_path, data)
      @xlsx_path = file_path
      @data = data
      @generic_storage = []
    end

    def self.parse_file(zip_file, file_path)
      (entry = zip_file.find_entry(RubyXL::from_root(file_path))) && self.new(file_path, entry.get_input_stream { |f| f.read })
    end

    def add_to_zip(zip_stream)
      return false if @data.nil?
      zip_stream.put_next_entry(RubyXL::from_root(self.xlsx_path))
      zip_stream.write(@data)
      true
    end
  end

  class PrinterSettingsFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
  end

  class CustomPropertyFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.customProperty'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty'
  end

  class DrawingFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'

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

  class BinaryImageFile < GenericStorageObject
    CONTENT_TYPE = 'image/jpeg'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
  end

  class VMLDrawingFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.vmlDrawing'
#    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
    REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'

    include RubyXL::RelationshipSupport

    define_relationship(RubyXL::BinaryImageFile)
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
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
  end

  class ControlPropertiesFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp'
  end

  class PivotCacheRecordsFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords'
  end

  class PivotCacheDefinitionFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition'

    include RubyXL::RelationshipSupport

    define_relationship(RubyXL::PivotCacheRecordsFile)
  end

  class PivotTableFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'

    include RubyXL::RelationshipSupport

    define_relationship(RubyXL::PivotCacheDefinitionFile)
  end

  class HyperlinkRelFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
  end

  class ThumbnailFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail'
    CONTENT_TYPE = 'image/x-wmf'
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

  class CustomXMLFile < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml'
  end

  class SlicerFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.ms-excel.slicer+xml'
    REL_TYPE     = 'http://schemas.microsoft.com/office/2007/relationships/slicer'
  end

  class SlicerCacheFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.ms-excel.slicerCache+xml'
    REL_TYPE     = 'http://schemas.microsoft.com/office/2007/relationships/slicerCache'
  end

  class OLEObjectFile < GenericStorageObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.oleObject'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject'
  end

  class ActiveX < GenericStorageObject
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/control'

    include RubyXL::RelationshipSupport

    def attach_relationship(rid, rf)
      case rf
      when RubyXL::ChartFile       then store_relationship(rf) # TODO
      else store_relationship(rf, :unknown)
      end
    end

  end

  class ActiveXBinary < GenericStorageObject
    REL_TYPE     = 'http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary'
  end

end
