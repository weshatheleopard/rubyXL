require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/simple_types'
#require 'rubyXL/objects/extensions'
require 'rubyXL/objects/relationships'

module RubyXL

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_from-1.html
  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_to-1.html
  class AnchorPoint < OOXMLObject
    define_child_node(RubyXL::IntegerNode, :node_name => 'xdr:col')
    define_child_node(RubyXL::IntegerNode, :node_name => 'xdr:colOff')
    define_child_node(RubyXL::IntegerNode, :node_name => 'xdr:row')
    define_child_node(RubyXL::IntegerNode, :node_name => 'xdr:rowOff')
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_ext-1.html
  class ShapeExtent < OOXMLObject
    define_attribute(:cx, :int, :required => true)
    define_attribute(:cy, :int, :required => true)
    define_element_name 'xdr:ext'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_grpSp-1.html
  class GroupShape < OOXMLObject
#    define_child_node(RubyXL::ShapeExtent, :node_name => 'xdr:ext')
#    define_child_node(RubyXL::Shape)
#    define_child_node(RubyXL::GroupShape)
#    define_child_node(RubyXL::CT_GraphicalObjectFrame)
#    define_child_node(RubyXL::ConnectionShape)
#    define_child_node(RubyXL::Picture)
    define_element_name 'xdr:grpSp'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_clientData-1.html
  class ClientData < OOXMLObject
    define_attribute(:fLocksWithSheet,  :bool, :default => true)
    define_attribute(:fPrintsWithSheet, :bool, :default => true)
    define_element_name 'xdr:clientData'
  end

  # http://www.datypic.com/sc/ooxml/t-draw-ssdraw_CT_OneCellAnchor.html
  class OneCellAnchor < OOXMLObject
    define_child_node(RubyXL::AnchorPoint, :node_name => 'xdr:from')
    define_child_node(RubyXL::ShapeExtent, :node_name => 'xdr:ext')
    # -- Choice [1..1] (EG_ObjectChoices)
    define_child_node(RubyXL::CT_Shape)
    define_child_node(RubyXL::GroupShape)
    define_child_node(RubyXL::CT_GraphicalObjectFrame)
    define_child_node(RubyXL::CT_Connector)
    define_child_node(RubyXL::CT_Picture)
    # --
    define_child_node(RubyXL::ClientData)
    define_element_name 'xdr:oneCellAnchor'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_twoCellAnchor-1.html
  class TwoCellAnchor < OOXMLObject
    define_child_node(RubyXL::AnchorPoint, :node_name => 'xdr:from')
    define_child_node(RubyXL::AnchorPoint, :node_name => 'xdr:to')
    # -- Choice [1..1] (EG_ObjectChoices)
    define_child_node(RubyXL::CT_Shape)
    define_child_node(RubyXL::GroupShape)
    define_child_node(RubyXL::CT_GraphicalObjectFrame)
    define_child_node(RubyXL::CT_Connector)
    define_child_node(RubyXL::CT_Picture)
    # --
    define_child_node(RubyXL::ClientData)
    define_attribute(:editAs, RubyXL::ST_EditAs, :default => 'twoCell')
    define_element_name 'xdr:twoCellAnchor'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_absoluteAnchor-1.html
  class AbsoluteAnchor < OOXMLObject
    define_child_node(RubyXL::CT_Point2D,  :node_name => 'xdr:pos')
    define_child_node(RubyXL::ShapeExtent, :node_name => 'xdr:ext')
    # -- Choice [1..1] (EG_ObjectChoices)
    define_child_node(RubyXL::CT_Shape)
    define_child_node(RubyXL::GroupShape)
    define_child_node(RubyXL::CT_GraphicalObjectFrame)
    define_child_node(RubyXL::CT_Connector)
    define_child_node(RubyXL::CT_Picture)
    # --
    define_child_node(RubyXL::ClientData)
    define_element_name 'xdr:absoluteAnchor'
  end

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
    define_child_node(RubyXL::AlternateContent)

    def attach_relationship(rid, rf)
      case rf
      when RubyXL::ChartFile       then store_relationship(rf) # TODO
      when RubyXL::BinaryImageFile then store_relationship(rf) # TODO
      else store_relationship(rf, :unknown)
      end
    end

    def xlsx_path
      ROOT.join('xl', 'drawings', "drawing#{file_index}.xml")
    end

  end

end



=begin

  # http://www.datypic.com/sc/ooxml/e-a_srcRect-1.html
  class SourceRectangle < OOXMLObject
    define_attribute(:l, :integer)
    define_attribute(:t, :integer)
    define_attribute(:r, :integer)
    define_attribute(:b, :integer)
    define_element_name 'a:srcRect'
  end

  # http://www.datypic.com/sc/ooxml/e-a_tile-1.html
  class Tile < OOXMLObject
    define_attribute(:tx, :int)
    define_attribute(:ty, :int)
    define_attribute(:sx, :int)
    define_attribute(:sy, :int)
    define_attribute(:flip, :string)
    define_attribute(:algn, :string)
  end

  # http://www.datypic.com/sc/ooxml/e-a_stretch-1.html
  class Stretch < OOXMLObject
  end

    def rel_id
      xdr_pic.xdr_blip_fill.a_blip.r_embed
    end

    def row
      xdr_from.xdr_row.value
    end

    def col
      xdr_from.xdr_col.value
    end

    def image_path
      @image_path
    end

    def image_path=(path)
      @image_path = path
    end
  end

#    def anchors
#      xdr_one_cell_anchor + xdr_two_cell_anchor + xdr_absolute_anchor
#    end
#
#    def set_image_paths
#      relationship_container.related_files.each do |rel_id, file|
#        anchor = anchors.select { |a| a.rel_id == rel_id }.first
#        anchor.image_path = file.xlsx_path
#      end
#    end

  end
end
=end
