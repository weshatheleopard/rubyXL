require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/extensions'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-a_ext-1.html
  class Extension < OOXMLObject
    define_attribute(:uri, :string)
    define_element_name 'a:ext'
    attr_accessor :raw_xml

    def self.parse(node)
      obj = new
      obj.raw_xml = node.to_xml
      obj
    end

    def write_xml(xml, node_name_override = nil)
      self.raw_xml
    end

  end

  class AExtensionStorageArea < OOXMLObject
    define_child_node(RubyXL::AExtension, :collection => true)
    define_element_name 'a:extLst'
  end

  class ColorScheme < OOXMLObject
    define_element_name 'a:clrScheme'
  end

  class FontScheme < OOXMLObject
    define_element_name 'a:fontScheme'
  end

  class FormatScheme < OOXMLObject
    define_element_name 'a:fmtScheme'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_themeElements-1.html
  class ThemeElements < OOXMLObject
    define_child_node(RubyXL::ColorScheme)
    define_child_node(RubyXL::FontScheme)
    define_child_node(RubyXL::FormatScheme)
    define_child_node(RubyXL::AExtensionStorageArea)
    define_element_name 'a:themeElements'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_theme.html
  class Theme < OOXMLObject
    define_attribute(:name, :string, :default => '')
    define_child_node(RubyXL::ThemeElements)
    define_child_node(RubyXL::ObjectDefaults)
#a:extraClrSchemeLst [0..1]    Extra Color Scheme List
#a:custClrLst [0..1]    Custom Color List
    define_child_node(RubyXL::AExtensionStorageArea)

    define_element_name 'a:theme'

    set_namespaces('xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main')
  end

end