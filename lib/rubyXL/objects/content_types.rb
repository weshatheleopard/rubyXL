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

    attr_accessor :owner

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

    def add_override(obj)
      return unless obj.class.const_defined?(:CONTENT_TYPE)
      overrides << RubyXL::ContentTypeOverride.new(:part_name => obj.xlsx_path, :content_type => obj.class::CONTENT_TYPE)
    end

    def before_write_xml
      self.defaults = []

      hsh = {}

      owner.rels_hash.each_pair { |klass, arr|
        puts "#{klass} -> #{arr.size} objects"

        arr.each { |x|

puts "file=#{x.xlsx_path}, ext=#{File.extname(x.xlsx_path)[1..-1]}"
puts x.class::CONTENT_TYPE if klass.const_defined?(:CONTENT_TYPE)

          ext = File.extname(x.xlsx_path)[1..-1]
          hsh[ext] ||= []
          hsh[ext] << klass::CONTENT_TYPE if klass.const_defined?(:CONTENT_TYPE)
        }
      }

      hsh.each_pair { |ext, content_types|
        next if ext.nil?
        content_types.uniq!

        if content_types.size == 1 then
          puts "UNIQUE: #{content_types.first}"
          defaults << RubyXL::ContentTypeDefault.new(:extension => ext, :content_type => content_types.first)
        end
      }

puts hsh.inspect

      defaults << RubyXL::ContentTypeDefault.new(:extension => 'xml', :content_type => 'application/xml' )

      # TODO: Need to only write these types when respective content is actually present.
      defaults << RubyXL::ContentTypeDefault.new(:extension => 'jpeg', :content_type => 'image/jpeg')
      defaults << RubyXL::ContentTypeDefault.new(:extension => 'png', :content_type => 'image/png')
      defaults << RubyXL::ContentTypeDefault.new(:extension => 'wmf', :content_type => 'image/x-wmf')

      true
    end

  end

end
