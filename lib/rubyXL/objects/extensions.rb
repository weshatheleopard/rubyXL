require 'rubyXL/objects/ooxml_object'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_ext-1.html
  class RawOOXML < OOXMLObject
    attr_accessor :raw_xml

    def self.parse(node, ignore)
      obj = new
      obj.raw_xml = node.to_xml
      obj
    end

    def write_xml(xml, node_name_override = nil)
      self.raw_xml
    end
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_ext-1.html
  class Extension < RawOOXML
    define_attribute(:uri, :string)
    define_element_name 'ext'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_extLst-1.html
  class ExtensionStorageArea < OOXMLObject
    define_child_node(RubyXL::Extension, :collection => true)
    define_element_name 'extLst'
  end

  class AlternateContent < RawOOXML
    define_element_name 'mc:AlternateContent'
  end

end
