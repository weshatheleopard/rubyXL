require 'pp'

module RubyXL

  # Parent class for defining OOXML based objects (not unlike Rails' +ActiveRecord+!)
  # Most importantly, provides functionality of parsing such objects from XML,
  # and marshalling them to XML.
  class OOXMLObject

    # Get the value of a [sub]class variable if it exists, or create the respective variable
    # with the passed-in +default+ (or +{}+, if not specified)
    # 
    # Throughout this class, we are setting class variables through explicit method calls
    # rather than by directly addressing the name of the variable because of context issues:
    # addressing variable by name creates it in the context of defining class, while calling
    # the setter/getter method addresses it in the context of descendant class, 
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

    # Defines an attribute of OOXML object.
    # === Parameters
    # * +attribute_name+ - Name of the element attribute as seen in the source XML. Can be either "String" or :Symbol
    #   * Special attibute name '_' (underscore) denotes the value of the element rather than attribute.
    # * +attribute_type+ - Specifies the conversion type for the attribute when parsing. Available options are:
    #   * :int - Integer
    #   * :float - Float
    #   * :string - String (no conversion)
    #   * :sqref - RubyXL::Sqref
    #   * :ref - RubyXL::Reference
    #   * :bool - Boolean ("1" and "true" convert to +true+, others to +false+)
    # * +extra_parameters+ - Hash of optional parameters as follows:
    #   * :accessor - Name of the accessor for this attribute to be defined on the object. If not provided, defaults to classidied +attribute_name+.
    #   * :default - Value this attribute defaults to if not explicitly provided.
    #   * :required - Whether this attribute is required when writing XML. If the value of the attrinute is not explicitly provided, +:default+ is written instead.
    #   * :values - List of acceptable values for this attribute (curently not used).
    # ==== Examples
    #   define_attribute(:outline, :bool, :default => true)
    # A Boolean attribute 'outline' with default value +true+ will be accessible by calling +obj.outline+
    #   define_attribute(:uniqueCount,  :int)
    # An Integer attribute 'uniqueCount' accessible as +obj.unique_count+
    #   define_attribute(:_,  :string, :accessor => :expression)
    # The value of the element will be accessible as a String by calling +obj.expression+
    #   define_attribute(:errorStyle, :string, :default => 'stop', :values => %w{ stop warning information })
    # A String attribute named 'errorStyle' will be accessible as +obj.error_style+, valid values are "stop", "warning", "information"
    def self.define_attribute(attr_name, attr_type, extra_params = {})
      attrs = obtain_class_variable(:@@ooxml_attributes)

      accessor = extra_params[:accessor] || accessorize(attr_name)
      attr_name = attr_name.to_s

      attrs[attr_name] = {
        :accessor   => accessor,
        :attr_type  => attr_type,
        :optional   => !extra_params[:required], 
        :default    => extra_params[:default],
        :valies     => extra_params[:values]
      }

      self.send(:attr_accessor, accessor)
    end
    
    def self.define_child_node(klass, extra_params = {})
      child_nodes = obtain_class_variable(:@@ooxml_child_nodes)
      child_node_name = (extra_params[:node_name] || klass.class_variable_get(:@@ooxml_tag_name)).to_s
      accessor = (extra_params[:accessor] || accessorize(child_node_name)).to_sym

      child_nodes[child_node_name] = { 
        :class => klass,
        :is_array => extra_params[:collection],
        :accessor => accessor
      }

      if extra_params[:collection] == :with_count then
        define_attribute(:count, :int, :required => true)
      end

      self.send(:attr_accessor, accessor)
    end

    def self.define_element_name(v)
      self.class_variable_set(:@@ooxml_tag_name, v)
    end

    def self.set_countable
      self.class_variable_set(:@@ooxml_countable, true)
      self.send(:attr_accessor, :count)
    end

    # Sets the list of namespaces on this object to be added when writing out XML. Valid only on top-level objects.
    # === Parameters
    # * +namespace_hash+ - Hash of namespaces in the form of <tt>"prefix" => "url"</tt>
    # ==== Examples
    #   set_namespaces('xmlns'   => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    #                  'xmlns:r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    def self.set_namespaces(namespace_hash)
      self.class_variable_set(:@@ooxml_namespaces, namespace_hash)
    end

    def write_xml(xml = nil, node_name_override = nil)
      if xml.nil? then
        seed_xml = Nokogiri::XML('<?xml version = "1.0" standalone ="yes"?>')
        seed_xml.encoding = 'UTF-8'
        result = self.write_xml(seed_xml)
        return result if result == ''
        seed_xml << result
        return seed_xml.to_xml({ :indent => 0, :save_with => Nokogiri::XML::Node::SaveOptions::AS_XML })
      end

      return '' unless before_write_xml

      attrs = obtain_class_variable(:@@ooxml_namespaces).dup

      obtain_class_variable(:@@ooxml_attributes).each_pair { |k, v|
        val = self.send(v[:accessor])

        if val.nil? then
          next if v[:optional]
          val = v[:default]
        end

        val = val &&
                case v[:attr_type]
                when :bool  then val ? '1' : '0'
                when :float then val.to_s.gsub(/\.0*$/, '') # Trim trailing zeroes
                else val
                end

        attrs[k] = val
      }

      element_text = attrs.delete('_')
      elem = xml.create_element(node_name_override || obtain_class_variable(:@@ooxml_tag_name), attrs, element_text)
      child_nodes = obtain_class_variable(:@@ooxml_child_nodes)
      child_nodes.each_pair { |child_node_name, child_node_params|
        obj = self.send(child_node_params[:accessor])
        unless obj.nil?
          if child_node_params[:is_array] then obj.each { |item| elem << item.write_xml(xml, child_node_name) unless item.nil? }
          else elem << obj.write_xml(xml, child_node_name)
          end
        end
      }
      elem
    end

    def initialize(params = {})
      obtain_class_variable(:@@ooxml_attributes).each_value { |v|
        instance_variable_set("@#{v[:accessor]}", params[v[:accessor]])
      }

      obtain_class_variable(:@@ooxml_child_nodes).each_value { |v|

        initial_value =
          if params.has_key?(v[:accessor]) then params[v[:accessor]]
          elsif v[:is_array] then []
          else nil
          end

        instance_variable_set("@#{v[:accessor]}", initial_value)
      }

      instance_variable_set("@count", 0) if obtain_class_variable(:@@ooxml_countable, false)
    end

    def self.process_attribute(obj, raw_value, params)
      val = raw_value &&
              case params[:attr_type]
              when :int    then Integer(raw_value)
              when :float  then Float(raw_value)
              when :string then raw_value
              when :sqref  then RubyXL::Sqref.new(raw_value)
              when :ref    then RubyXL::Reference.new(raw_value)
              when :bool   then ['1', 'true'].include?(raw_value)
              end              
      obj.send("#{params[:accessor]}=", val)
    end
    private_class_method :process_attribute

    def self.parse(node)
      node = Nokogiri::XML.parse(node) if node.is_a?(IO) || node.is_a?(String)

      if node.is_a?(Nokogiri::XML::Document) then
#        @namespaces = node.namespaces
        node = node.root
#        ignorable_attr = node.attributes['Ignorable']
#        @ignorables << ignorable_attr.value if ignorable_attr
      end

      obj = self.new

      known_attributes = obtain_class_variable(:@@ooxml_attributes)

      content_params = known_attributes['_']
      process_attribute(obj, node.text, content_params) if content_params

      node.attributes.each_pair { |attr_name, attr|
        attr_name = if attr.namespace then "#{attr.namespace.prefix}:#{attr.name}"
                    else attr.name
                    end

        attr_params = known_attributes[attr_name]

        next if attr_params.nil?
        # raise "Unknown attribute: #{attr_name}" if attr_params.nil?
        process_attribute(obj, attr.value, attr_params)
      }

      known_child_nodes = obtain_class_variable(:@@ooxml_child_nodes)

      unless known_child_nodes.empty?
        node.element_children.each { |child_node|

          child_node_name = if child_node.namespace.prefix then
                              "#{child_node.namespace.prefix}:#{child_node.name}"
                            else child_node.name 
                            end

          child_node_params = known_child_nodes[child_node_name]
          raise "Unknown child node: #{child_node_name}" if child_node_params.nil?
          parsed_object = child_node_params[:class].parse(child_node)
          if child_node_params[:is_array] then
            index = parsed_object.index_in_collection
            collection = obj.send(child_node_params[:accessor])
            if index.nil? then
              collection << parsed_object
            else
              collection[index] = parsed_object
            end
          else
            obj.send("#{child_node_params[:accessor]}=", parsed_object)
          end
        }
      end

      obj
    end

    def dup
      new_copy = super
      new_copy.count = 0 if obtain_class_variable(:@@ooxml_countable, false)
      new_copy
    end

    def index_in_collection
      nil
    end

    # Subclass provided filter to perform last-minute operations (cleanup, count, etc.) immediately prior to write,
    # along with option to terminate the actual write if +false+ is returned (for example, to avoid writing
    # the collection's root node if the collection is empty).
    def before_write_xml
      child_nodes = obtain_class_variable(:@@ooxml_child_nodes)
      child_nodes.each_pair { |child_node_name, child_node_params|
        self.count = self.send(child_node_params[:accessor]).size if child_node_params[:is_array] == :with_count
      }
      true 
    end

    def add_to_zip(zipfile)
      xml_string = write_xml
      return if xml_string.empty?
      zipfile.get_output_stream(self.class.filepath) { |f| f << xml_string }
    end

    def self.filepath
      raise 'Subclass responsebility'
    end

    def self.parse_file(dirpath)
      full_path = File.join(dirpath, filepath)
      return nil unless File.exist?(full_path)
      parse(File.open(full_path, 'r'))
    end

    private
    def self.accessorize(str)
      acc = str.to_s.dup
      acc.gsub!(/([A-Z\d]+)([A-Z][a-z])/,'\1_\2')
      acc.gsub!(/([a-z\d])([A-Z])/,'\1_\2')
      acc.gsub!(':','_')
      acc.downcase.to_sym
    end

  end
end
