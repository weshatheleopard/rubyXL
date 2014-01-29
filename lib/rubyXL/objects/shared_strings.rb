require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/text'
require 'rubyXL/objects/extensions'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sst.html
  class SharedStringsTable < OOXMLObject
    define_attribute(:uniqueCount,  :int)
    define_child_node(RubyXL::RichText, :collection => :with_count, :node_name => 'si')
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'sst'
  end

end
