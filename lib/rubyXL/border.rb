module RubyXL

  class Border
    attr_accessor :count, :edges

    def initialize(args = {})
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


  class BorderEdge
    attr_accessor :style, :color

    def initialize(attrs = {})
      @style = attrs['style']
      @color = attrs['color']
    end

    def self.parse(xml)
      return nil if xml.attributes.empty? && xml.element_children.empty?

      edge = self.new
      edge.style = RubyXL::Parser.attr_string(xml, 'style')

      xml.element_children.each { |node|
        case node.name
        when 'color' then edge.color = RubyXL::Color.parse(node)
        else raise "node name #{node.name} not supported yet"
        end
      }

      edge
    end

    def build_xml(xml)
      xml.fill {
        xml.patternFill(:patternType => pattern_type) {
          fg_color && fg_color.build_xml(xml, 'fgColor')
          bg_color && bg_color.build_xml(xml, 'bgColor')
        }
      }
    end

  end

end
