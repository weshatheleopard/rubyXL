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
      zipfile.get_output_stream(self.xlsx_path) { |f| f << @data }
    end
  end

  class PrinterSettings < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings'
    end
  end

  class Drawing < GenericStorageObject
    attr_accessor :relationship_container

    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'
    end

    def self.content_type
      'application/vnd.openxmlformats-officedocument.drawing+xml'
    end

    def load_relationships(dir_path, base_file_name)

      self.relationship_container = RubyXL::DrawingRelationships.load_relationship_file(dir_path, base_file_name)

      if relationship_container then
        relationship_container.load_related_files(dir_path, base_file_name)

        related_files = relationship_container.related_files
        related_files.each_pair { |rid, rf|
          case rf
when 1 then 1
#          when RubyXL::PrinterSettings then self.generic_storage << rf # TODO
#          when RubyXL::Comments        then self.generic_storage << rf # TODO
#          when RubyXL::VMLDrawing      then self.generic_storage << rf # TODO
#          when RubyXL::Drawing         then self.generic_storage << rf # TODO
          else
            self.generic_storage << rf
puts "!!>DEBUG: unattached: #{rf.class}"
          end
        }
      end
    end

    include RubyXL::RelationshipSupport

    def related_objects
      relationship_container.owner = self
      [ relationship_container ] + generic_storage
    end

  end

  class Chart < GenericStorageObject
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
          when RubyXL::ChartColors then self.generic_storage << rf # TODO
          when RubyXL::ChartStyle  then self.generic_storage << rf # TODO
          else
            self.generic_storage << rf
puts "!!>DEBUG: unattached: #{rf.class}"
          end
        }
      end
    end

    include RubyXL::RelationshipSupport

    def related_objects
      relationship_container.owner = self
      [ relationship_container ] + generic_storage
    end

  end

  class VMLDrawing < GenericStorageObject
    attr_accessor :relationship_container

    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing'
    end

#    def self.content_type
#      'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
#    end

    def load_relationships(dir_path, base_file_name)

      self.relationship_container = RubyXL::DrawingRelationships.load_relationship_file(dir_path, base_file_name)

      if relationship_container then
        relationship_container.load_related_files(dir_path, base_file_name)

        related_files = relationship_container.related_files
        related_files.each_pair { |rid, rf|
          case rf
when 1 then 1
#          when RubyXL::ChartColors then self.generic_storage << rf # TODO
#          when RubyXL::ChartStyle  then self.generic_storage << rf # TODO
          else
            self.generic_storage << rf
puts "!!>DEBUG: unattached: #{rf.class}"
          end
        }
      end
    end

    def related_objects
      generic_storage
    end

  end

  class ChartColors < GenericStorageObject
    def self.rel_type
      'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
    end

    def self.content_type
      'application/vnd.ms-office.chartcolorstyle+xml'
    end
  end

  class ChartStyle < GenericStorageObject
    def self.rel_type
      'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
    end

    def self.content_type
      'application/vnd.ms-office.chartstyle+xml'
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

  class BinaryImage < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
    end
  end

  class HyperlinkRel < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
    end
  end

  class Thumbmail < GenericStorageObject
    def self.rel_type
      'http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail'
    end
  end

end
