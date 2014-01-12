module RubyXL

  class BooleanValue < OOXMLObject
    define_attribute(:val, :bool, :required => true, :default => true)
  end

  class StringValue < OOXMLObject
    define_attribute(:val, :string, :required => true)
  end

  class IntegerValue < OOXMLObject
    define_attribute(:val, :int, :required => true)
  end

  class FloatValue < OOXMLObject
    define_attribute(:val, :float, :required => true)
  end


  # http://www.schemacentral.com/sc/ooxml/e-ssml_font-1.html
  class Font < OOXMLObject
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

    attr_accessor :count

    def initialize(attrs = {})
      super
      @count = 0
    end

    def dup
      new = super
      new.count = 1
      new
    end

    def build_xml(xml)
      xml.font {
        xml.sz(:val => sz && sz.val)
        xml.b if b && b.val
        xml.i if i && i.val
        xml.u if u && u.val
        xml.strike if strike && strike.val
        @color.build_xml(xml) if @color
        xml.family(:val => family.val) if family && family.val
        xml.scheme(:val => scheme.val) if scheme && scheme.val
        xml.name(:val => name.val) if name && name.val
      }
    end

    def ==(other)
     (!(self.i && self.i.val) == !(other.i && other.i.val)) &&
       (!(self.b && self.b.val) == !(other.b && other.b.val)) &&
       (!(self.u && self.u.val) == !(other.u && other.u.val)) &&
       (!(self.strike && self.strike.val) == !(other.strike && other.strike.val)) &&
       ((self.name && self.name.val) == (other.name && other.name.val)) &&
       ((self.sz && self.sz.val) == (other.sz && other.sz.val)) &&
       (self.color == other.color) # Need to write proper comparison for color
    end

    def is_italic
      i && i.val
    end

    def set_italic(val)
      self.i ||= RubyXL::BooleanValue.new
      self.i.val = val
    end

    def is_bold
      b && b.val
    end

    def set_bold(val)
      self.b ||= RubyXL::BooleanValue.new
      self.b.val = val
    end

    def is_underlined
      u && u.val
    end

    def set_underline(val)
      self.u ||= RubyXL::BooleanValue.new
      self.u.val = val
    end

    def is_strikethrough
      strike && strike.val
    end

    def set_strikethrough(val)
      self.strike ||= RubyXL::BooleanValue.new
      self.strike.val = val
    end

    def get_name
      name && name.val
    end

    def set_name(val)
      self.name ||= RubyXL::StringValue.new
      self.name.val = val
    end

    def get_size
      sz && sz.val
    end

    def set_size(val)
      self.sz ||= RubyXL::FloatValue.new
      self.sz.val = val
    end

    def get_rgb_color
      color && color.rgb
    end

    # Helper method to modify the font color
    def set_rgb_color(font_color)
      self.color = RubyXL::Color.new(:rgb => font_color.to_s)
    end

  end

end
