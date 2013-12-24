module RubyXL

  class Fill
    attr_accessor :count

    def initialize(args = {})
      @count = 0
    end

    def self.parse(xml)
      node = xml.first_element_child
      case node.name
      when 'patternFill'  then RubyXL::PatternFill.parse(node)
      when 'gradientFill' then RubyXL::GradientFill.parse(node)
      else raise 'Unknown fill type'
      end
    end 

  end


  class PatternFill < Fill
    attr_accessor :pattern_type, :fg_color, :bg_color

    def initialize(attrs = {})
      @pattern_type = attrs['pattern_type']
      @fg_color     = attrs['fg_color']
      @fb_color     = attrs['bg_color']
      super
    end

    def self.parse(xml)
      fill = self.new
      fill.pattern_type = xml.attributes['patternType'].value

      xml.element_children.each { |node|
        case node.name
        when 'fgColor' then fill.fg_color = RubyXL::Color.parse(node)
        when 'bgColor' then fill.bg_color = RubyXL::Color.parse(node)
        else raise "node name #{node.name} not supported yet"
        end
      }

      fill
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


  class GradientFill < Fill
    attr_accessor :degree, :stop0_color, :stop1_color

    def self.parse(xml)
      fill = self.new
      fill.degree = xml.attributes['degree'].value

      xml.element_children.each { |node|
        case node.name
        when 'stop' then
          case node.attributes['position'].value
          when '0' then fill.stop0_color = RubyXL::Color.parse(node.first_element_child)
          when '1' then fill.stop1_color = RubyXL::Color.parse(node.first_element_child)
          else raise "stop position #{node.attributes['position'].value} not supported yet"
          end
        else raise "node name #{node.name} not supported yet"
        end
      }

      fill
    end 

    def build_xml(xml)
      xml.fill {
        xml.gradientFill(:degree => degree) {
          if stop0_color then
            xml.stop(:position => 0) { stop0_color.build_xml(xml) }
          end

          if stop1_color then
            xml.stop(:position => 1) { stop1_color.build_xml(xml) }
          end
        }
      }
    end

  end

end
