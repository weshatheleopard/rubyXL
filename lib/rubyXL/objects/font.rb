require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/container_nodes'
require 'rubyXL/objects/color'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_font-1.html
  class Font < OOXMLObject
    # Since we have no capability to load the actual fonts, we'll have to live with the default.
    MAX_DIGIT_WIDTH = 7 # Calibri 11 pt @ 96 dpi

    define_child_node(RubyXL::StringValue,  :node_name => :name)
    define_child_node(RubyXL::IntegerValue, :node_name => :charset)
    define_child_node(RubyXL::IntegerValue, :node_name => :family)
    define_child_node(RubyXL::BooleanValue, :node_name => :b)
    define_child_node(RubyXL::BooleanValue, :node_name => :i)
    define_child_node(RubyXL::BooleanValue, :node_name => :strike)
    define_child_node(RubyXL::BooleanValue, :node_name => :outline)
    define_child_node(RubyXL::BooleanValue, :node_name => :shadow)
    define_child_node(RubyXL::BooleanValue, :node_name => :condense)
    define_child_node(RubyXL::BooleanValue, :node_name => :extend)
    define_child_node(RubyXL::Color)
    define_child_node(RubyXL::FloatValue,   :node_name => :sz)
    define_child_node(RubyXL::BooleanValue, :node_name => :u)
    define_child_node(RubyXL::StringValue,  :node_name => :vertAlign)
    define_child_node(RubyXL::StringValue,  :node_name => :scheme)
    define_element_name 'font'

    def is_italic
      i && (i.val || true)
    end

    def set_italic(val)
      self.i = RubyXL::BooleanValue.new(:val => val)
    end

    def is_bold
      b && (b.val || true)
    end

    def set_bold(val)
      self.b = RubyXL::BooleanValue.new(:val => val)
    end

    def is_underlined
      u && (u.val || true)
    end

    def set_underline(val)
      self.u = RubyXL::BooleanValue.new(:val => val)
    end

    def is_strikethrough
      strike && (strike.val || true)
    end

    def set_strikethrough(val)
      self.strike = RubyXL::BooleanValue.new(:val => val)
    end

    def get_name
      name && name.val
    end

    def set_name(val)
      self.name = RubyXL::StringValue.new(:val => val)
    end

    def get_size
      sz && sz.val
    end

    def set_size(val)
      self.sz = RubyXL::FloatValue.new(:val => val)
    end

    def get_rgb_color
      color && color.rgb
    end

    # Helper method to modify the font color
    def set_rgb_color(font_color)
      self.color = RubyXL::Color.new(:rgb => font_color.to_s)
    end

    def self.default(size = 10)
      self.new(:name => RubyXL::StringValue.new(:val => 'Verdana'),
               :sz   => RubyXL::FloatValue.new(:val => size) )
    end

  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_fonts-1.html
  class Fonts < OOXMLContainerObject
    define_child_node(RubyXL::Font, :collection => :with_count)
    define_element_name 'fonts'

    def self.default
      self.new(:_ => [ RubyXL::Font.default(10), RubyXL::Font.default(8) ])
    end
  end

end
