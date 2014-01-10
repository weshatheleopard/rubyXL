module RubyXL
  # http://www.schemacentral.com/sc/ooxml/e-ssml_dataValidation-1.html
  class Formula < OOXMLObject
    define_attribute(:_,                 :string, :accessor => :expression)
    define_element_name 'f'
  end

end
