module RubyXL

  # Eventually, the entire code for Worksheet will be moved here. One small step at a time!

  # http://www.schemacentral.com/sc/ooxml/e-ssml_legacyDrawing-1.html
  class LegacyDrawing < OOXMLObject
    define_attribute(:'r:id',            :string, :required => true)
    define_element_name 'legacyDrawing'
  end

end
