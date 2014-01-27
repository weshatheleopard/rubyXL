module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_gradientFill-1.html
  class Stop < OOXMLObject
    define_attribute(:position, :float)
    define_child_node(RubyXL::Color)
    define_element_name 'stop'
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
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_fill-1.html
  class Fill < OOXMLObject
    define_child_node(RubyXL::PatternFill)
    define_child_node(RubyXL::GradientFill)
    define_element_name 'fill'
    set_countable
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_fills-1.html
  class FillContainer < OOXMLObject
    define_child_node(RubyXL::Fill, :collection => :with_count, :accessor => :fills)
    define_element_name 'fills'

    def self.defaults
      self.new(:fills => [
                 RubyXL::Fill.new(:pattern_fill => RubyXL::PatternFill.new(:pattern_type => 'none')),
                 RubyXL::Fill.new(:pattern_fill => RubyXL::PatternFill.new(:pattern_type => 'gray125'))
               ])
    end
  end

end
