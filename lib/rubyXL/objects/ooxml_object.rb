require 'pp'

module RubyXL
  class OOXMLObject

    # Throughout this class, setting class variables through explicit method calls rather
    # than by directly addressing the name of the variable because of context issues:
    # addressing variable by name creates it in the context of defining class,
    # while calling the setter/getter method addresses it in the context of descendant class, 
    # which is what we need.
    def self.obtain_class_variable(var_name, default = {})
      if class_variable_defined?(var_name) then 
        self.class_variable_get(var_name)
      else
        self.class_variable_set(var_name, default)
      end
    end

    def obtain_class_variable(var_name, default = {})
      self.class.obtain_class_variable(var_name, default)
    end
    private :obtain_class_variable

    def self.define_attribute(attr_name, attr_type, extra_params = {})
      attrs = obtain_class_variable(:@@ooxml_attributes)

      accessor = extra_params[:accessor]
      if accessor.nil? then 
        accessor = attr_name.to_s
        accessor.gsub!(/([A-Z\d]+)([A-Z][a-z])/,'\1_\2')
        accessor.gsub!(/([a-z\d])([A-Z])/,'\1_\2')
        accessor.downcase!
        accessor = accessor.to_sym
      end

      attrs[accessor] = {
        :attr_name  => attr_name.to_s,
        :attr_type  => attr_type,
        :optional   => !extra_params[:required], 
        :default    => extra_params[:default],
        :validation => extra_params[:values]
      }

      self.send(:attr_accessor, accessor)
    end
    
    def self.define_child_node(klass, extra_params = {})
      child_nodes = obtain_class_variable(:@@ooxml_child_nodes)
      child_node_name = (extra_params[:node_name] || klass.class_variable_get(:@@ooxml_tag_name)).to_s
      accessor = (extra_params[:accessor] || child_node_name).to_sym

      child_nodes[child_node_name] = { 
        :class => klass,
        :is_array => extra_params[:collection],
        :accessor => accessor
      }

      self.send(:attr_accessor, accessor)
    end

    def self.define_element_name(v)
      self.class_variable_set(:@@ooxml_tag_name, v)
    end

    def write_xml(xml, node_name_override = nil)
      before_write_xml if self.respond_to?(:before_write_xml)
      attrs = prepare_attributes
      element_text = attrs.delete('_')
      elem = xml.create_element(node_name_override || obtain_class_variable(:@@ooxml_tag_name), attrs, element_text)
      child_nodes = obtain_class_variable(:@@ooxml_child_nodes)
      child_nodes.each_pair { |child_node_name, child_node_params|
        obj = self.send(child_node_params[:accessor])
        unless obj.nil?
          if child_node_params[:is_array] then obj.each { |item| elem << item.write_xml(xml) }
          else elem << obj.write_xml(xml, child_node_name)
          end
        end
      }
      elem
    end

    def initialize(params = {})
      return super unless self.class.class_variable_defined?(:@@ooxml_attributes)
      obtain_class_variable(:@@ooxml_attributes).each_key { |k| instance_variable_set("@#{k}", params[k]) }
      obtain_class_variable(:@@ooxml_child_nodes).each_pair { |k, v|

        initial_value =
          if params.has_key?(k) then params[k]
          elsif v[:is_array] then []
          else nil
          end

        instance_variable_set("@#{v[:accessor]}", initial_value)
      }
    end

    def self.parse(node)
      obj = self.new

      obtain_class_variable(:@@ooxml_attributes).each_pair { |k, v|

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

      known_child_nodes = obtain_class_variable(:@@ooxml_child_nodes)

      unless known_child_nodes.empty?
        node.element_children.each { |child_node|
          child_node_name = child_node.name
          child_node_params = known_child_nodes[child_node_name]
          raise "Unknown child node: #{child_node_name}" if child_node_params.nil?
          parsed_object = child_node_params[:class].parse(child_node)
          if child_node_params[:is_array] then obj.send(child_node_params[:accessor]) << parsed_object
          else obj.send("#{child_node_params[:accessor]}=", parsed_object)
          end
        }
      end

      obj
    end

    def prepare_attributes
      xml_attrs = {}

      obtain_class_variable(:@@ooxml_attributes).each_pair { |k, v|
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
