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
    define_attribute(:horizontal,        :horizontal,      :string, false, nil,
                       %w{general left center right fill justify centerContinuous distributed})
    define_attribute(:vertical,          :vertical,        :string, false, nil,
                       %w{top center bottom justify distributed})
    define_attribute(:text_rotation,     :textRotation,    :int)
    define_attribute(:wrap_text,         :wrapText,        :bool)
    define_attribute(:indent,            :indent,          :int)
    define_attribute(:relative_indent,   :relativeIndent,  :int)
    define_attribute(:justify_last_line, :justifyLastLine, :bool)
    define_attribute(:shrink_to_fit,     :shrinkToFit,     :bool)
    define_attribute(:reading_order,     :readingOrder,    :int)
    define_element_name 'alignment'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_protection-1.html
  class Protection < OOXMLObject
    define_attribute(:locked, :locked, :bool)
    define_attribute(:hidden, :hidden, :bool)
    define_element_name 'protection'
  end


end
