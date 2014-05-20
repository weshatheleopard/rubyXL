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
    attr_accessor :workbook

    define_child_node(RubyXL::ContentTypeDefault,  :collection => true, :accessor => :defaults)
    define_child_node(RubyXL::ContentTypeOverride, :collection => true, :accessor => :overrides)

    set_namespaces('http://schemas.openxmlformats.org/package/2006/content-types' => '')
    define_element_name 'Types'

    def self.xlsx_path
      '[Content_Types].xml'
    end

    def self.save_order
      999  # Must be saved last, so it has time to accumulate overrides from all others.
    end

    def add_override(obj)
      return unless obj.class.respond_to?(:content_type)
      overrides << RubyXL::ContentTypeOverride.new(:part_name => "/#{obj.xlsx_path}", :content_type => obj.class.content_type)
    end

    def before_write_xml
      self.defaults = []
      if @workbook.rels_hash[RubyXL::PrinterSettings] then
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
      defaults << RubyXL::ContentTypeDefault.new(:extension => 'vml', :content_type => 'application/vnd.openxmlformats-officedocument.vmlDrawing')

=begin
      workbook.charts.each_pair { |k, v|
        case k
        when /^chart\d*.xml$/ then
          overrides << RubyXL::ContentTypeOverride.new(:part_name => "/#{@workbook.charts.local_dir_path}/#{k}",
                         :content_type => 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
        when /^style\d*.xml$/ then
          overrides << RubyXL::ContentTypeOverride.new(:part_name => "/#{@workbook.charts.local_dir_path}/#{k}",
                         :content_type => 'application/vnd.ms-office.chartstyle+xml')
        when /^colors\d*.xml$/ then
          overrides << RubyXL::ContentTypeOverride.new(:part_name => "/#{@workbook.charts.local_dir_path}/#{k}",
                         :content_type => 'application/vnd.ms-office.chartcolorstyle+xml')
        end
      }
=end
=begin
      workbook.drawings.each_pair { |k, v|
        case k
        when /^drawing\d*.xml$/ then
          overrides << RubyXL::ContentTypeOverride.new(:part_name => "/#{@workbook.drawings.local_dir_path}/#{k}",
                         :content_type => 'application/vnd.openxmlformats-officedocument.drawing+xml')
        when /^vmlDrawing\d*.vml$/ then 
          # more proper fix: <Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>
          overrides << RubyXL::ContentTypeOverride.new(:part_name => "/#{@workbook.drawings.local_dir_path}/#{k}",
                         :content_type => 'application/vnd.openxmlformats-officedocument.vmlDrawing')
        end
      }
=end
=begin
      unless workbook.external_links.nil?
        1.upto(workbook.external_links.size - 1) { |i|
          overrides << RubyXL::ContentTypeOverride.new(:part_name => "/xl/externalLinks/externalLink#{i}.xml",
                       :content_type => 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml')
        }
      end

      unless @workbook.macros.empty?
        overrides << RubyXL::ContentTypeOverride.new(:part_name => '/xl/vbaProject.bin',
                       :content_type => 'application/vnd.ms-office.vbaProject')
      end
=end

      true
    end

  end

end
