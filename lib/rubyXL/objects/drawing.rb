require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/simple_types'
#require 'rubyXL/objects/extensions'
require 'rubyXL/objects/relationships'

module RubyXL

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_col-1.html
  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_row-1.html
  class Point < OOXMLObject
    define_attribute(:_, :int, :accessor => :value)
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_colOff-1.html
  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_rowOff-1.html
  class PointOffset < OOXMLObject
    define_attribute(:_, :int, :accessor => :value)
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_from-1.html
  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_to-1.html
  class AnchorPoint < OOXMLObject
    define_child_node(RubyXL::Point, node_name: 'xdr:col')
    define_child_node(RubyXL::PointOffset, node_name: 'xdr:colOff')
    define_child_node(RubyXL::Point, node_name: 'xdr:row')
    define_child_node(RubyXL::PointOffset, node_name: 'xdr:rowOff')
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_ext-1.html
  class ShapeExtent < OOXMLObject
    define_attribute(:cx, :int)
    define_attribute(:cy, :int)
    define_element_name 'xdr:ext'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_sp-1.html
  class Shape < OOXMLObject
    define_attribute(:macro, :string)
    define_attribute(:textlink, :string)
    define_attribute(:fLocksText, :bool)
    define_attribute(:fPublished, :bool)
    define_element_name 'xdr:sp'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_grpSp-1.html
  class GroupShape < OOXMLObject
    define_element_name 'xdr:grpSp'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_graphicFrame-1.html
  class GraphicFrame < OOXMLObject
    define_element_name 'xdr:graphicFrame'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_cxnSp-1.html
  class ConnectionShape < OOXMLObject
    define_element_name 'xdr:cxnSp'
  end

  # http://www.datypic.com/sc/ooxml/e-a_blip-1.html
  class Blip < OOXMLObject
    define_attribute('r:embed', :string)
    define_element_name 'a:blip'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_cNvPr-2.html
  class ConnectionNonVisualProperties < OOXMLObject
    define_element_name 'xdr:cNvPr'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_cNvPicPr-1.html
  class NonVisualPictureDrawingProperties < OOXMLObject
    define_element_name 'xdr:cNvPicPr'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_nvPicPr-1.html
  class NonVisualPictureProperties < OOXMLObject
    define_child_node(RubyXL::ConnectionNonVisualProperties)
    define_child_node(RubyXL::NonVisualPictureDrawingProperties)
    define_element_name 'xdr:nvPicPr'
  end

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

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_blipFill-1.html
  class BlipFill < OOXMLObject
    define_child_node(RubyXL::Blip)
    define_child_node(RubyXL::SourceRectangle)
    define_child_node(RubyXL::Tile, node_name: 'a:tile')
    define_child_node(RubyXL::Stretch, node_name: 'a:stretch')
    define_attribute(:dpi, :integer)
    define_attribute(:rotWithShape, :integer)
    define_element_name 'xdr:blipFill'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_spPr-1.html
  class ShapeProperties < OOXMLObject
    define_element_name 'xdr:spPr'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_style-1.html
  class ShapeStyle < OOXMLObject
    define_element_name 'xdr:style'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_pic-1.html
  class Picture < OOXMLObject
    define_child_node(RubyXL::NonVisualPictureProperties)
    define_child_node(RubyXL::BlipFill)
    define_child_node(RubyXL::ShapeProperties)
    define_child_node(RubyXL::ShapeStyle)
    define_element_name 'xdr:pic'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_clientData-1.html
  class ClientData < OOXMLObject
    define_element_name 'xdr:clientData'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_oneCellAnchor-1.html
  class OneCellAnchor < OOXMLObject
    define_child_node(RubyXL::AnchorPoint, :node_name => :from)
    define_child_node(RubyXL::ShapeExtent, :node_name => :ext)
    define_child_node(RubyXL::Shape)
    define_child_node(RubyXL::GroupShape)
    define_child_node(RubyXL::GraphicFrame)
    define_child_node(RubyXL::ConnectionShape)
    define_child_node(RubyXL::Picture)
    define_child_node(RubyXL::ClientData)
    define_element_name 'xdr:oneCellAnchor'
  end

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_twoCellAnchor-1.html
  class TwoCellAnchor < OOXMLObject
    define_child_node(RubyXL::AnchorPoint, :node_name => 'xdr:from')
    define_child_node(RubyXL::AnchorPoint, :node_name => 'xdr:to')
    define_child_node(RubyXL::Shape)
    define_child_node(RubyXL::GroupShape)
    define_child_node(RubyXL::GraphicFrame)
    define_child_node(RubyXL::ConnectionShape)
    define_child_node(RubyXL::Picture)
    define_child_node(RubyXL::ClientData)
    define_attribute(:editAs, RubyXL::ST_EditAs, :default => 'twoCell')
    define_element_name 'xdr:twoCellAnchor'

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

  # http://www.datypic.com/sc/ooxml/e-draw-ssdraw_absoluteAnchor-1.html
  class AbsoluteAnchor < OOXMLObject
    define_child_node(RubyXL::AnchorPoint, :node_name => :pos)
    define_child_node(RubyXL::ShapeExtent, :node_name => :ext)
    define_child_node(RubyXL::Shape)
    define_child_node(RubyXL::GroupShape)
    define_child_node(RubyXL::GraphicFrame)
    define_child_node(RubyXL::ConnectionShape)
    define_child_node(RubyXL::Picture)
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

    def anchors
      xdr_one_cell_anchor + xdr_two_cell_anchor + xdr_absolute_anchor
    end

    def set_image_paths
      relationship_container.related_files.each do |rel_id, file|
        anchor = anchors.select { |a| a.rel_id == rel_id }.first
        anchor.image_path = file.xlsx_path
      end
    end
  end
end
