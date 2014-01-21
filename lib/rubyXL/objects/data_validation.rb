module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_dataValidation-1.html
  class ValidationFormula < OOXMLObject
    define_attribute(:_,                 :string, :accessor => :expression)
    define_element_name 'f'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_dataValidation-1.html
  class DataValidation < OOXMLObject
    define_attribute(:type,             :string, :default => 'none',
                        :values => %w{ none whole decimal list date time textLength custom })
    define_attribute(:errorStyle,       :string, :default => 'stop',
                        :values => %w{ stop warning information })
    define_attribute(:imeMode,          :string, :default => 'noControl',
                        :values => %w{ noControl off on disabled hiragana fullKatakana halfKatakana
                            fullAlpha halfAlpha fullHangul halfHangul })
    define_attribute(:operator,         :string, :default => 'between',
                        :values => %w{ between notBetween equal notEqual lessThan lessThanOrEqual
                            greaterThan greaterThanOrEqual })
    define_attribute(:allowBlank,       :bool, :default => false)
    define_attribute(:showDropDown,     :bool, :default => false)
    define_attribute(:showInputMessage, :bool, :default => false)
    define_attribute(:showErrorMessage, :bool, :default => false)
    define_attribute(:errorTitle,       :string)
    define_attribute(:error,            :string)
    define_attribute(:promptTitle,      :string)
    define_attribute(:prompt,           :string)
    define_attribute(:sqref,            :sqref, :required => true)

    define_child_node(RubyXL::ValidationFormula, :node_name => :formula1)
    define_child_node(RubyXL::ValidationFormula, :node_name => :formula2)
    define_element_name 'dataValidation'

  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_dataValidations-1.html
  class DataValidations < OOXMLObject
    define_attribute(:disablePrompts, :bool, :default => false)
    define_attribute(:xWindow,        :int)
    define_attribute(:yWindow,        :int)
    define_child_node(RubyXL::DataValidation, :collection => :with_count)
    define_element_name 'dataValidations'
  end

end
