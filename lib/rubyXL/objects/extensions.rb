module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_ext-1.html
  class Extension < OOXMLObject
    define_attribute(:uri, :string)
    define_element_name 'ext'
    attr_accessor :raw_xml

    def self.parse(node)
      obj = new
      obj.raw_xml = node.to_xml
      obj
    end

    def write_xml(xml, node_name_override = nil)
      xml << self.raw_xml
    end

  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_extLst-1.html
  class ExtensionStorageArea < OOXMLObject
    define_child_node(RubyXL::Extension, :collection => true)
    define_element_name 'extLst'
  end

end