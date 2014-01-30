require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/text'
require 'rubyXL/objects/extensions'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sst.html
  class SharedStringsTable < OOXMLObject
    # According to http://msdn.microsoft.com/en-us/library/office/gg278314.aspx,
    # +count+ and +uniqueCount+ may be either both missing, or both present. Need to validate.
    define_attribute(:uniqueCount,  :int)
    define_child_node(RubyXL::RichText, :collection => :with_count, :node_name => 'si', :accessor => :strings)
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'sst'
    set_namespaces('xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    def initialize(*params)
      super
      # So far, going by the structure that the original creator had in mind. However, 
      # since the actual implementation is now extracted into a separate class, 
      # we will be able to transparrently change it later if needs be.
      @index_by_content = {}
    end

    def [](index)
      strings[index]
    end

    def empty?
      strings.empty?
    end

    def add(str, index = nil)
      index ||= strings.size
      strings[index] = RubyXL::Text.new(:value => str)
      @index_by_content[str] = index
    end

    def get_index(str, add_if_missing = false)
      index = @index_by_content[str]
      index = add(str) if index.nil? && add_if_missing
      index 
    end

  end

end
