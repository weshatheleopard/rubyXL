module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_cellStyle-1.html
  class CellStyle < OOXMLObject
    define_attribute(:name,           :name,          :string)
    define_attribute(:xf_id,          :xfId,          :int,    :required)
    define_attribute(:builtin_id,     :builtinId,     :int)
    define_attribute(:i_level,        :iLevel,        :int)
    define_attribute(:hidden,         :hidden,        :bool)
    define_attribute(:custom_builtin, :customBuiltin, :bool)
    define_element_name 'cellStyle'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_xf-1.html
  class XF < OOXMLObject

    define_attribute(:num_fmt_id,          :numFmtId,          :int)
    define_attribute(:font_id,             :fontId,            :int)
    define_attribute(:fill_id,             :fillId,            :int)
    define_attribute(:border_id,           :borderId,          :int)
    define_attribute(:xf_id,               :xfId,              :int)
    define_attribute(:quote_prefix,        :quotePrefix,       :bool, false)
    define_attribute(:pivot_button,        :pivotButton,       :bool, false)
    define_attribute(:apply_number_format, :applyNumberFormat, :bool)
    define_attribute(:apply_font,          :applyFont,         :bool)
    define_attribute(:apply_fill,          :applyFill,         :bool)
    define_attribute(:apply_border,        :applyBorder,       :bool)
    define_attribute(:apply_alignment,     :applyAlignment,    :bool)
    define_attribute(:apply_protection,    :applyProtection,   :bool)
    define_element_name 'xf'

    def self.parse(node)
      xf = super

      node.element_children.each { |child_node|
        case child_node.name                 
        when 'alignment'  then xf.alignment  = RubyXL::Alignment.parse(child_node)
        when 'protection' then xf.protection = RubyXL::Protection.parse(child_node)
        else raise "Node type #{child_node.name} not implemented"
        end
      }

      xf
    end 

    def write_xml
      super # TODO
    end
=begin

<xf numFmtId="14" fontId="60" fillId="11" borderId="22" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
<alignment horizontal="left"/>
<protection locked="0"/>
</xf>
=end
  end


  # http://www.schemacentral.com/sc/ooxml/e-ssml_alignment-1.html
  class Alignment < OOXMLObject
=begin
numFmtId	[0..1]	ssml:ST_NumFmtId	Number Format Id	
fontId	[0..1]	ssml:ST_FontId	Font Id	
fillId	[0..1]	ssml:ST_FillId	Fill Id	
borderId	[0..1]	ssml:ST_BorderId	Border Id	
xfId	[0..1]	ssml:ST_CellStyleXfId	Format Id	
quotePrefix	[0..1]	xsd:boolean	Quote Prefix	Default value is "false".
pivotButton	[0..1]	xsd:boolean	Pivot Button	Default value is "false".
applyNumberFormat	[0..1]	xsd:boolean	Apply Number Format	
applyFont	[0..1]	xsd:boolean	Apply Font	
applyFill	[0..1]	xsd:boolean	Apply Fill	
applyBorder	[0..1]	xsd:boolean	Apply Border	
applyAlignment	[0..1]	xsd:boolean	Apply Alignment	
applyProtection	[0..1]	xsd:boolean	Apply Protection
=end
    define_element_name 'alignment'

  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_protection-1.html
  class Protection < OOXMLObject
    define_attribute(:locked, :locked, :bool)
    define_attribute(:hidden, :hidden, :bool)
    define_element_name 'protection'
  end


end
