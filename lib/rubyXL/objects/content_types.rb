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

    set_namespaces(:xmlns => 'http://schemas.openxmlformats.org/package/2006/content-types')
    define_element_name 'Types'

    def self.xlsx_path
      '[Content_Types].xml'
    end

    def generate_override(obj)
      RubyXL::ContentTypeOverride.new(:part_name => "/#{obj.xlsx_path}", :content_type => obj.class.content_type)
    end

    def before_write_xml
      self.defaults = []
      unless @workbook.printer_settings.empty?
        defaults << RubyXL::ContentTypeDefault.new(:extension => 'bin',
                      :content_type => 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings')
      end

      defaults << RubyXL::ContentTypeDefault.new(:extension => 'rels',
                      :content_type => 'application/vnd.openxmlformats-package.relationships+xml' )

      defaults << RubyXL::ContentTypeDefault.new(:extension => 'xml',
                      :content_type => 'application/xml' )

      defaults << RubyXL::ContentTypeDefault.new(:extension => 'jpeg', :content_type => 'image/jpeg' )

      self.overrides = []
      overrides << generate_override(workbook)
      workbook.worksheets.each { |sheet| overrides << generate_override(sheet) }
      overrides << generate_override(workbook.stylesheet)
      overrides << generate_override(workbook.document_properties)
      overrides << generate_override(workbook.core_properties)
      overrides << generate_override(workbook.shared_strings_container) unless workbook.shared_strings_container.empty?
      overrides << generate_override(workbook.calculation_chain) unless workbook.calculation_chain.nil?
#      overrides << generate_override(workbook.theme)

      overrides << RubyXL::ContentTypeOverride.new(:part_name => '/xl/theme/theme1.xml',
                      :content_type => 'application/vnd.openxmlformats-officedocument.theme+xml')

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

      workbook.drawings.each_pair { |k, v|
        overrides << RubyXL::ContentTypeOverride.new(:part_name => "/#{@workbook.drawings.local_dir_path}/#{k}",
                       :content_type => 'application/vnd.openxmlformats-officedocument.drawing+xml')
#      ('Extension'=>'vml', 'ContentType'=>'application/vnd.openxmlformats-officedocument.vmlDrawing')
      }

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

      true
    end

  end


=begin

#           if @workbook.macros.nil? && @workbook.drawings.empty?
            root << xml.create_element('Override', {
                      :PartName    => '/xl/workbook.xml',
                      :ContentType => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml' })
#           else
#             xml.Override('PartName'=>'/xl/workbook.xml',
#               'ContentType'=>"application/vnd.ms-excel.sheet.macroEnabled.main+xml")
#           end



          })



<Default Extension="jpeg" ContentType="image/jpeg"/>
<Override PartName="/xl/chartsheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>
<Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
<Override PartName="/xl/charts/style2.xml" ContentType="application/vnd.ms-office.chartstyle+xml"/>
<Override PartName="/xl/charts/colors2.xml" ContentType="application/vnd.ms-office.chartcolorstyle+xml"/>
=end

end
