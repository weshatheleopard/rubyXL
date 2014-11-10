require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/simple_types'

module RubyXL

  class BorderEdge < OOXMLObject
    define_attribute(:style, RubyXL::ST_BorderStyle, :default => 'none')
    define_child_node(RubyXL::Color)
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_border-2.html
  class Border < OOXMLObject
    define_attribute(:diagonalUp,   :bool)
    define_attribute(:diagonalDown, :bool)
    define_attribute(:outline,      :bool, :default => true)
    define_child_node(RubyXL::BorderEdge, :node_name => :left)
    define_child_node(RubyXL::BorderEdge, :node_name => :right)
    define_child_node(RubyXL::BorderEdge, :node_name => :top)
    define_child_node(RubyXL::BorderEdge, :node_name => :bottom)
    define_child_node(RubyXL::BorderEdge, :node_name => :diagonal)
    define_child_node(RubyXL::BorderEdge, :node_name => :vertical)
    define_child_node(RubyXL::BorderEdge, :node_name => :horizontal)
    define_element_name 'border'

    def get_edge_style(direction)
      edge = self.send(direction)
      edge && edge.style
    end

    def set_edge_style(direction, style)
      self.send("#{direction}=", RubyXL::BorderEdge.new(:style => style))
    end
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_borders-1.html
  class Borders < OOXMLContainerObject
    define_child_node(RubyXL::Border, :collection => :with_count)
    define_element_name 'borders'

    def self.default
      self.new(:_ => [ RubyXL::Border.new ])
    end

  end

end
