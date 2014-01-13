module RubyXL

  class BorderEdge < OOXMLObject
    define_attribute(:style,   :string)
    define_child_node(RubyXL::Color, :default => 'none', :values => 
                        %w{ none thin medium dashed dotted thick double hair
                            mediumDashed dashDot mediumDashDot dashDotDot slantDashDot } )
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

    attr_accessor :count

    def initialize(args = {})
      super
      @count = 0
    end

    def get_edge_style(direction)
      edge = self.send(direction)
      edge && edge.style
    end

    def set_edge_style(direction, style)
      self.send("#{direction}=", RubyXL::BorderEdge.new(:style => style))
    end

  end

end
