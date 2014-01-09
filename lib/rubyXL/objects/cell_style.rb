module RubyXL
  # http://www.schemacentral.com/sc/ooxml/e-ssml_cellStyle-1.html

  class CellStyle < OOXMLObject
    define_attribute(:name,           :name,          :string, true)
    define_attribute(:xf_id,          :xfId,          :int)
    define_attribute(:builtin_id,     :builtinId,     :int,    true)
    define_attribute(:i_level,        :iLevel,        :int,    true)
    define_attribute(:hidden,         :hidden,        :bool,   true)
    define_attribute(:custom_builtin, :customBuiltin, :bool,   true)
    define_element_name 'cellStyle'
  end

end
