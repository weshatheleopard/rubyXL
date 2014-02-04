require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/extensions'

module RubyXL
  # http://www.schemacentral.com/sc/ooxml/e-ssml_c-1.html
  class CalculationChainCell < OOXMLObject
    define_attribute(:r, :ref,  :accessor => :ref)
    define_attribute(:i, :int,  :accessor => :sheet_id,    :default => 0)
    define_attribute(:s, :bool, :accessor => :child_chain, :default => false)
    define_attribute(:l, :bool, :accessor => :new_dep_lvl, :default => false)
    define_attribute(:t, :bool, :accessor => :new_thread,  :default => false)
    define_attribute(:a, :bool, :accessor => :array,       :default => false)
    define_element_name 'c'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_calcChain.html
  class CalculationChain < OOXMLObject
    define_child_node(RubyXL::CalculationChainCell, :collection => true, :accessor => :cells)
    define_child_node(RubyXL::ExtensionStorageArea)

    define_element_name 'calcChain'
    set_namespaces('xmlns'     => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    def self.filepath
      File.join('xl', 'calcChain.xml')
    end
  end
end
