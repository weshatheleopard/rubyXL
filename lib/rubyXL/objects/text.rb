module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_t-1.html
  class Text < OOXMLObject
    define_attribute(:_, :string, :accessor => :value)
    define_element_name 't'
  end

  class RichTextRun < OOXMLObject
#    define_child_node(RubyXL::RunProperties)      # rPr
    define_child_node(RubyXL::Text)               # t
    define_element_name 'r'
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
    define_attribute(:type,      :string, :default => 'fullwidthKatakana',
                        :values => %w{ halfwidthKatakana fullwidthKatakana Hiragana noConversion })
    define_attribute(:alignment, :string, :default => 'left',
                        :values => %w{ noControl left center distributed })
    define_element_name 'phoneticPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_is-1.html
  class RichText < OOXMLObject
    define_child_node(RubyXL::Text)
    define_child_node(RubyXL::RichTextRun)
    define_child_node(RubyXL::PhoneticRun)
    define_child_node(RubyXL::PhoneticProperties) # phoneticPr
    define_element_name 'is'
  end

end
