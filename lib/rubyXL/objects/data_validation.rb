module RubyXL
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
    define_attribute(:allowBlank,       :bool, :default => 'false')
    define_attribute(:showDropDown,     :bool, :default => 'false')
    define_attribute(:showInputMessage, :bool, :default => 'false')
    define_attribute(:showErrorMessage, :bool, :default => 'false')
    define_attribute(:errorTitle,       :string)
    define_attribute(:error,            :string)
    define_attribute(:promptTitle,      :string)
    define_attribute(:prompt,           :string)
    define_attribute(:sqref,            :sqref, :required => true)

    attr_accessor :formula1, :formula2

    def initialize
      @formula1 = @formula2 = nil
      super
    end

    def self.parse(node)
      val = super

      node.element_children.each { |child_node|
        case child_node.name
        when 'formula1' then val.formula1 = child_node.text
        when 'formula2' then val.formula2 = child_node.text
        else raise "Node type #{child_node.name} not implemented"
        end
      }

      val
    end 

    def write_xml(xml)
      node = xml.create_element('dataValidation', prepare_attributes)
      node << xml.create_element('formula1', @formula1) unless formula1.nil?
      node << xml.create_element('formula2', @formula2) unless formula2.nil?
      node
    end

  end

end
