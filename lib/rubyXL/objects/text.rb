require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/simple_types'
require 'rubyXL/objects/container_nodes'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_t-1.html
  class Text < OOXMLObject
    define_attribute(:_,           :string, :accessor => :value)
    define_attribute(:'xml:space', :string)
    define_element_name 't'

    def before_write_xml
      self.xml_space = (value.is_a?(String) && ((value =~ /\A\s/) || (value =~ /\s\Z/) || value.include?("\n"))) ? 'preserve' : nil
      true
    end

    def to_s
      value.to_s
    end
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_rPr-1.html
  class RunProperties < OOXMLObject
    define_child_node(RubyXL::StringValue,  :node_name => :rFont)
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
    define_element_name 'rPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_r-2.html
  class RichTextRun < OOXMLObject
    define_child_node(RubyXL::RunProperties)
    define_child_node(RubyXL::Text)
    define_element_name 'r'

    def to_s
      t.to_s
    end

  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_rPh-1.html
  class PhoneticRun < OOXMLObject
    define_attribute(:sb, :int, :required => true)
    define_attribute(:eb, :int, :required => true)
    define_child_node(RubyXL::Text)
    define_element_name 'rPh'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_phoneticPr-1.html
  class PhoneticProperties < OOXMLObject
    define_attribute(:fontId,    :int,    :required => true)
    define_attribute(:type,      RubyXL::ST_PhoneticType,      :default => 'fullwidthKatakana')
    define_attribute(:alignment, RubyXL::ST_PhoneticAlignment, :default => 'left')
    define_element_name 'phoneticPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_is-1.html
  class RichText < OOXMLObject
    define_child_node(RubyXL::Text)
    define_child_node(RubyXL::RichTextRun, :collection => true)
    define_child_node(RubyXL::PhoneticRun, :collection => true)
    define_child_node(RubyXL::PhoneticProperties)
    define_element_name 'is'

    def to_s
      str = t.to_s
      r && r.each { |rtr| str << rtr.to_s if rtr }
      str
    end
  end

end
