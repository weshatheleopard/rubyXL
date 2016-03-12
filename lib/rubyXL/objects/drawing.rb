require 'rubyXL/objects/ooxml_object'
#require 'rubyXL/objects/simple_types'
#require 'rubyXL/objects/extensions'
#require 'rubyXL/objects/relationships'
#require 'rubyXL/objects/sheet_common'

module RubyXL

  # http://www.datypic.com/sc/ooxml/s-dml-spreadsheetDrawing.xsd.html
  class DrawingFile < OOXMLTopLevelObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawing+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing'

    include RubyXL::RelationshipSupport

    define_relationship(RubyXL::BinaryImageFile)

    define_element_name 'xdr:wsDr'

    set_namespaces('http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing' => 'xdr',
                   'http://schemas.openxmlformats.org/drawingml/2006/main' => 'a', 
                   'http://schemas.openxmlformats.org/officeDocument/2006/relationships' => 'r')

    define_child_node(RubyXL::OneCellAnchor,  :collection => [0..-1])
    define_child_node(RubyXL::TwoCellAnchor,  :collection => [0..-1])
    define_child_node(RubyXL::AbsoluteAnchor, :collection => [0..-1])

    def attach_relationship(rid, rf)
      case rf
      when RubyXL::ChartFile       then store_relationship(rf) # TODO
      when RubyXL::BinaryImageFile then store_relationship(rf) # TODO
      else store_relationship(rf, :unknown)
      end
    end

    def xlsx_path
      ROOT.join('xl', 'drawings', "drawings#{file_index}.xml")
    end

  end

end
