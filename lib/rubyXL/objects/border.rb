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

    attr_accessor :count, :edges

    def initialize(args = {})
#      super
      @count = 0
      @edges = {}
    end

    def self.parse(xml)
      border = self.new

      xml.element_children.each { |side_node|
        side = side_node.name
        case side
        when 'left', 'right', 'top', 'bottom', 'diagonal' then
          border.edges[side] = RubyXL::BorderEdge.parse(side_node)
        else raise "Unknown border side: #{side_node.name}"
        end
      }
      border
    end 

    def build_xml(xml)
      xml.border {
        build_edge(xml, 'left')
        build_edge(xml, 'right')
        build_edge(xml, 'top')
        build_edge(xml, 'bottom')
        build_edge(xml, 'diagonal')
      }
    end

    def build_edge(xml, side)
      edge_obj = @edges[side]
      return xml.send(side.to_sym) if edge_obj.nil?

      xml.send(side.to_sym, { :style => edge_obj.style }) {
        edge_obj.color.build_xml(xml) unless edge_obj.color.nil?
      }
    end
    private :build_edge

  end

end
