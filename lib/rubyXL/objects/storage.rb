module RubyXL

  class GenericStorageObject

    attr_accessor :xlsx_path, :data, :workbook

    def initialize
      @workbook = nil
      @xlsx_path = nil
      @data = nil
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
      zipfile.get_output_stream(self.xlsx_path) { |f| f << @data }
    end
  end

  class PrinterSettings < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
    end
  end

  class Drawing < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
    end

    def self.content_type
      'application/vnd.openxmlformats-officedocument.drawing+xml'
    end
  end

  class VMLDrawing < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
    end
  end

  class Table < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table'
    end
  end

  class ControlProperties < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp'
    end
  end

  class PivotTable < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable'
    end
  end

end