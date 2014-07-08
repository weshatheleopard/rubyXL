require 'rubyXL/objects/ooxml_object'

module RubyXL

  class ContentTypeDefault < OOXMLObject
    define_attribute(:Extension,   :string)
    define_attribute(:ContentType, :string)
    define_element_name 'Default'
  end

  class ContentTypeOverride < OOXMLObject
    define_attribute(:PartName,   :string)
    define_attribute(:ContentType, :string)
    define_element_name 'Override'
  end

  class ContentTypes < OOXMLTopLevelObject
    SAVE_ORDER = 999  # Must be saved last, so it has time to accumulate overrides from all others.

    define_child_node(RubyXL::ContentTypeDefault,  :collection => true, :accessor => :defaults)
    define_child_node(RubyXL::ContentTypeOverride, :collection => true, :accessor => :overrides)

    set_namespaces('http://schemas.openxmlformats.org/package/2006/content-types' => '')
    define_element_name 'Types'

    def self.xlsx_path
      ROOT.join('[Content_Types].xml')
    end

    def xlsx_path
      self.class.xlsx_path
    end

    def before_write_xml
      content_types_by_ext = {}

      # Collect all extensions and corresponding content types
      root.rels_hash.each_pair { |klass, objects|
        objects.each { |obj|
          next unless klass.const_defined?(:CONTENT_TYPE)
          ext = obj.xlsx_path.extname[1..-1]
          next if ext.nil?
          content_types_by_ext[ext] ||= []
          content_types_by_ext[ext] << klass::CONTENT_TYPE 
        }
      }

      self.defaults = [ RubyXL::ContentTypeDefault.new(:extension => 'xml', :content_type => 'application/xml') ]

      # Determine which content types are used most often, and add them to the list of defaults
      content_types_by_ext.each_pair { |ext, content_types_arr|
        next if ext.nil? || defaults.any? { |d| d.extension == ext } 

        counter = {}  # Count the number of occurrences of a given +content_type+.
        content_types_arr.each { |ct| counter[ct] = (counter[ct] || 0) + 1 }
        most_used_content_type = counter.to_a.max_by{ |pair| pair.last }.first

        defaults << RubyXL::ContentTypeDefault.new(:extension => ext,
                                                   :content_type => most_used_content_type)
      }

      self.overrides = []

      # Add overrides for the files with known extensions but different content types.
      root.rels_hash.each_pair { |klass, objects|
        objects.each { |obj|
          next unless defined?(klass::CONTENT_TYPE)
          ext = obj.xlsx_path.extname[1..-1]
          next if ext.nil?
          next if defaults.any? { |d| (d.content_type == klass::CONTENT_TYPE) && (d.extension == ext) }
          overrides << RubyXL::ContentTypeOverride.new(:part_name => obj.xlsx_path,
                                                       :content_type => klass::CONTENT_TYPE)
        }
      }

      true
    end

  end

end
