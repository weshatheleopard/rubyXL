module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_numFmt-1.html
  class NumberFormat < OOXMLObject
    define_attribute(:numFmtId,   :int,    :required => true)
    define_attribute(:formatCode, :string, :required => true)
    define_element_name 'numFmt'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_numFmts-1.html
  class NumberFormatContainer < OOXMLObject
    define_child_node(RubyXL::NumberFormat, :collection => :with_count, :accessor => :number_formats)
    define_element_name 'numFmts'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_cellStyleXfs-1.html
  class CellStyleXFContainer < OOXMLObject
    define_child_node(RubyXL::XF, :collection => :with_count, :accessor => :cell_style_xfs)
    define_element_name 'cellStyleXfs'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_cellXfs-1.html
  class CellXFs < OOXMLObject
    define_child_node(RubyXL::XF, :collection => :with_count, :accessor => :xfs)
    define_element_name 'cellXfs'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_cellStyles-1.html
  class CellStyleContainer < OOXMLObject
    define_child_node(RubyXL::CellStyle, :collection => :with_count, :accessor => :cell_styles)
    define_element_name 'cellStyles'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_dxf-1.html
  class DXF < OOXMLObject
    define_child_node(RubyXL::Font)
    define_child_node(RubyXL::NumberFormat)
    define_child_node(RubyXL::Fill)
    define_child_node(RubyXL::Alignment)
    define_child_node(RubyXL::Border)
    define_child_node(RubyXL::Protection)
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'dxf'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_dxfs-1.html
  class DXFs < OOXMLObject
    define_child_node(RubyXL::DXF, :collection => :with_count)
    define_element_name 'dxfs'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_tableStyle-1.html
  class TableStyle < OOXMLObject
    define_attribute(:name,  :string, :required => true)
    define_attribute(:pivot, :bool,   :default => true)
    define_attribute(:table, :bool,   :default => true)
    define_attribute(:count, :int)
    define_element_name 'tableStyle'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_tableStyles-1.html
  class TableStyles < OOXMLObject
    define_attribute(:defaultTableStyle, :string)
    define_attribute(:defaultPivotStyle, :string)
    define_child_node(RubyXL::TableStyle, :collection => :with_count)
    define_element_name 'tableStyles'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_colors-1.html
  class Colors < OOXMLObject
    define_child_node(RubyXL::Color, :collection => :true, :node_name => 'indexedColors')
    define_child_node(RubyXL::Color, :collection => :true, :node_name => 'mruColors')
    define_element_name 'colors'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_styleSheet.html
  class Stylesheet < OOXMLObject
    define_child_node(RubyXL::NumberFormatContainer, :accessor => :number_format_container)
    define_child_node(RubyXL::FontContainer,         :accessor => :font_container)
    define_child_node(RubyXL::FillContainer,         :accessor => :fill_container)
    define_child_node(RubyXL::BorderContainer,       :accessor => :border_container)
    define_child_node(RubyXL::CellStyleXFContainer,  :accessor => :cell_style_xf_container)
    define_child_node(RubyXL::CellXFs)
    define_child_node(RubyXL::CellStyleContainer,    :accessor => :cell_style_container)
    define_child_node(RubyXL::DXFs)
    define_child_node(RubyXL::TableStyles)
    define_child_node(RubyXL::Colors)
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'styleSheet'
    set_namespaces('xmlns'       => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                   'xmlns:r'     => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                   'xmlns:mc'    => 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                   'xmlns:x14ac' => 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
                   'xmlns:mv'    => 'urn:schemas-microsoft-com:mac:vml')

    def initialize(*args)
      super
      @format_hash = nil
    end

    def self.default
      self.new(:cell_xfs => RubyXL::CellXFs.new(
                 :xfs => [ 
                   RubyXL::XF.new(
                     :num_fmt_id => 0, :font_id => 0, :fill_id => 0, :border_id => 0, :xfId => 0
                   )
                 ]
               ),
               :font_container => RubyXL::FontContainer.new(
                 :fonts => [ 
                   RubyXL::Font.new(:name => RubyXL::StringValue.new(:val => 'Verdana'),
                                    :sz => RubyXL::FloatValue.new(:val => 10) ),
                   RubyXL::Font.new(:name => RubyXL::StringValue.new(:val => 'Verdana'),
                                    :sz => RubyXL::FloatValue.new(:val => 8) )
                 ]
               ),
               :fill_container => RubyXL::FillContainer.new(
                 :fills => [
                   RubyXL::Fill.new(:pattern_fill => RubyXL::PatternFill.new(:pattern_type => 'none')),
                   RubyXL::Fill.new(:pattern_fill => RubyXL::PatternFill.new(:pattern_type => 'gray125'))
                 ]
               ),
               :border_container => RubyXL::BorderContainer.new(
                 :borders => [ RubyXL::Border.new ]
               ),
               :cell_style_container => RubyXL::CellStyleContainer.new(
                 :cell_styles => [ RubyXL::CellStyle.new(:builtin_id => 0, :name => 'Normal', :xf_id => 0) ]
               ),
               :cell_style_xf_container => RubyXL::CellStyleXFContainer.new(
                 :cell_style_xfs => [ RubyXL::XF.new(:num_fmt_id => 0, :font_id => 0, :fill_id => 0, :border_id => 0) ]
               )
             )
    end

    def number_format(format_id)
      if @format_hash.nil? then
        @format_hash = {}
        if number_format_container then
          number_format_container.number_formats.each { |fmt| @format_hash[fmt.num_fmt_id] = fmt }
        end
      end

      @format_hash[format_id]
    end

  end

end
