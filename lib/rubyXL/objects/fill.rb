module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_gradientFill-1.html
  class Stop < OOXMLObject
    define_attribute(:position, :float)
    define_child_node(RubyXL::Color)
    define_element_name 'stop'

    def build_xml(xml)
      xml.stop(:position => position) {
        color && color.build_xml(xml)
      }
    end
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_patternFill-1.html
  class PatternFill < OOXMLObject
    define_attribute(:patternType, :string, :values =>
                       %w{ none solid mediumGray darkGray lightGray
                           darkHorizontal darkVertical darkDown darkUp darkGrid darkTrellis
                           lightHorizontal lightVertical lightDown lightUp lightGrid lightTrellis
                           gray125 gray0625 })
    define_child_node(RubyXL::Color, :node_name => :fgColor )
    define_child_node(RubyXL::Color, :node_name => :bgColor )
    define_element_name 'patternFill'

    def build_xml(xml)
      xml.patternFill(:patternType => pattern_type) {
        fg_color && fg_color.build_xml(xml, 'fgColor')
        bg_color && bg_color.build_xml(xml, 'bgColor')
      }
    end

  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_gradientFill-1.html
  class GradientFill < OOXMLObject
    define_attribute(:type,   :string, :values => %w{ linear path }, :default => 'linear')
    define_attribute(:degree, :float,  :default => 0)
    define_attribute(:left,   :float,  :default => 0)
    define_attribute(:right,  :float,  :default => 0)
    define_attribute(:top,    :float,  :default => 0)
    define_attribute(:bottom, :float,  :default => 0)
    define_child_node(RubyXL::Stop, :collection => true)
    define_element_name 'gradientFill'

    def build_xml(xml)
      xml.gradientFill(:degree => degree) {
        stop.each { |s|
          s.build_xml(xml)
        }
      }
    end

  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_fill-1.html
  class Fill < OOXMLObject
    attr_accessor :count

    define_child_node(RubyXL::PatternFill)
    define_child_node(RubyXL::GradientFill)
    define_element_name 'fill'

    def initialize(*args)
      super
      @count = 0
    end

    def build_xml(xml)
      xml.fill {
        pattern_fill && pattern_fill.build_xml(xml)
        gradient_fill && gradient_fill.build_xml(xml)
      }
    end

  end


end
