module RubyXL
  class OOXMLObject

    # Throughout this class, setting class variables through explicit method calls rather
    # than by directly addressing the name of the variable because of context issues:
    # addressing variable by name creates it in the context of defining class,
    # while calling the setter/getter method addresses it in the context of descendant class, 
    # which is what we need.

    def self.define_attribute(accessor, attr_name, attr_type, optional = false, default = nil, value_list = nil)

      if class_variable_defined?(:@@ooxml_attributes) then
        attrs = self.class_variable_get(:@@ooxml_attributes)
      else
        attrs = self.class_variable_set(:@@ooxml_attributes, {})
      end

      params = {
        :attr_name  => attr_name.to_s,
        :attr_type  => attr_type,
        :optional   => optional, 
        :default    => default,
        :validation => value_list,
      }

      attrs[accessor] = params

      self.send(:attr_accessor, accessor)
    end

    def self.define_element_name(v)
      self.class_variable_set(:@@ooxml_tag_name, v)
    end

    def write_xml(xml)
      before_write_xml if self.respond_to?(:before_write_xml)
      attrs = prepare_attributes
      content = attrs.delete('_')
      xml.create_element(self.class.class_variable_get(:@@ooxml_tag_name), attrs, content)
    end

    def initialize(params = {})
      return super unless self.class.class_variable_defined?(:@@ooxml_attributes)
      attrs = self.class.class_variable_get(:@@ooxml_attributes)
      attrs.each_key { |k| instance_variable_set("@#{k}", params[k]) }
    end

    def self.parse(node)
      obj = self.new

      self.class_variable_get(:@@ooxml_attributes).each_pair { |k, v|

        raw_value = if v[:attr_name] == '_' then node.text
                    else
                      attr = node.attributes[v[:attr_name]]
                      attr && attr.value
                    end
                    
        val = raw_value &&
                case v[:attr_type]
                when :int    then Integer(raw_value)
                when :float  then Float(raw_value)
                when :string then raw_value
                when :sqref  then RubyXL::Sqref.new(raw_value)
                when :ref    then RubyXL::Reference.new(raw_value)
                when :bool   then raw_value.to_i == 1
                end              


        obj.send("#{k}=", val)
      }

      obj
    end

    def prepare_attributes
      xml_attrs = {}

      self.class.class_variable_get(:@@ooxml_attributes).each_pair { |k, v|
        val = self.send(k)

        if val.nil? then
          next if v[:optional]
          val = v[:default]
        end

        val = val &&
                case v[:attr_type]
                when :bool then val ? '1' : '0'
                else val
                end

        xml_attrs[v[:attr_name]] = val
      }

      xml_attrs
    end
  end
end