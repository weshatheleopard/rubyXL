module RubyXL
  # http://www.schemacentral.com/sc/ooxml/e-ssml_dataValidation-1.html
  class DataValidation < OOXMLObject
    define_attribute(:type,               :type,             :string, true,
                        %w{ whole decimal list date time textLength custom })
    define_attribute(:error_style,        :errorStyle,       :string, true,
                        %w{ stop warning information })
    define_attribute(:ime_mode,           :imeMode,          :string, true,
                        %w{ noControl off on disabled hiragana fullKatakana halfKatakana
                            fullAlpha halfAlpha fullHangul halfHangul })
    define_attribute(:operator,           :operator,         :string, true,
                        %w{ between notBetween equal notEqual lessThan lessThanOrEqual
                            greaterThan greaterThanOrEqual })
    define_attribute(:allow_blank,        :allowBlank,       :int,    true)
    define_attribute(:show_drop_down,     :showDropDown,     :int,    true)
    define_attribute(:show_input_message, :showInputMessage, :int,    true)
    define_attribute(:show_error_message, :showErrorMessage, :int,    true)
    define_attribute(:error_title,        :errorTitle,       :string, true)
    define_attribute(:error,              :error,            :string, true)
    define_attribute(:prompt_title,       :promptTitle,      :string, true)
    define_attribute(:prompt,             :prompt,           :string, true)
    define_attribute(:sqref,              :sqref,            :sqref)

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
