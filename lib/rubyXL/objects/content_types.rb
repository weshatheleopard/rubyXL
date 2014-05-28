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
    attr_accessor :owner

    define_child_node(RubyXL::ContentTypeDefault,  :collection => true, :accessor => :defaults)
    define_child_node(RubyXL::ContentTypeOverride, :collection => true, :accessor => :overrides)

    set_namespaces('http://schemas.openxmlformats.org/package/2006/content-types' => '')
    define_element_name 'Types'

    def self.xlsx_path
      '[Content_Types].xml'
    end

    def xlsx_path
      self.class.xlsx_path
    end

    def self.save_order
      999  # Must be saved last, so it has time to accumulate overrides from all others.
    end

    def add_override(obj)
      return unless obj.class.const_defined?(:CONTENT_TYPE)
      overrides << RubyXL::ContentTypeOverride.new(:part_name => "/#{obj.xlsx_path}", :content_type => obj.class::CONTENT_TYPE)
    end

    def before_write_xml
      self.defaults = []
      if owner.rels_hash[RubyXL::PrinterSettingsFile] then
        defaults << RubyXL::ContentTypeDefault.new(:extension => 'bin',
                      :content_type => 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings')
      end

      defaults << RubyXL::ContentTypeDefault.new(:extension => 'rels',
                      :content_type => 'application/vnd.openxmlformats-package.relationships+xml' )

      defaults << RubyXL::ContentTypeDefault.new(:extension => 'xml',
                      :content_type => 'application/xml' )

      # TODO: Need to only write these types when respective content is actually present.
      defaults << RubyXL::ContentTypeDefault.new(:extension => 'jpeg', :content_type => 'image/jpeg')
      defaults << RubyXL::ContentTypeDefault.new(:extension => 'png', :content_type => 'image/png')
      defaults << RubyXL::ContentTypeDefault.new(:extension => 'wmf', :content_type => 'image/x-wmf')
      defaults << RubyXL::ContentTypeDefault.new(:extension => 'vml', :content_type => 'application/vnd.openxmlformats-officedocument.vmlDrawing')

      true
    end

  end

end
