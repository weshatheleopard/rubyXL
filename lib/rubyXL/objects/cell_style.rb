require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/simple_types'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_cellStyle-1.html
  class CellStyle < OOXMLObject
    define_attribute(:name,          :string)
    define_attribute(:xfId,          :int,  :required => true)
    define_attribute(:builtinId,     :int)
    define_attribute(:iLevel,        :int)
    define_attribute(:hidden,        :bool)
    define_attribute(:customBuiltin, :bool)
    define_element_name 'cellStyle'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_alignment-1.html
  class Alignment < OOXMLObject
    define_attribute(:horizontal,      RubyXL::ST_HorizontalAlignment)
    define_attribute(:vertical,        RubyXL::ST_VerticalAlignment)
    define_attribute(:textRotation,    :int)
    define_attribute(:wrapText,        :bool)
    define_attribute(:indent,          :int)
    define_attribute(:relativeIndent,  :int)
    define_attribute(:justifyLastLine, :bool)
    define_attribute(:shrinkToFit,     :bool)
    define_attribute(:readingOrder,    :int)
    define_element_name 'alignment'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_protection-1.html
  class Protection < OOXMLObject
    define_attribute(:locked, :bool)
    define_attribute(:hidden, :bool)
    define_element_name 'protection'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_xf-1.html
  class XF < OOXMLObject
    define_attribute(:numFmtId,          :int)
    define_attribute(:fontId,            :int)
    define_attribute(:fillId,            :int)
    define_attribute(:borderId,          :int)
    define_attribute(:xfId,              :int)
    define_attribute(:quotePrefix,       :bool, :default => false )
    define_attribute(:pivotButton,       :bool, :default => false )
    define_attribute(:applyNumberFormat, :bool)
    define_attribute(:applyFont,         :bool)
    define_attribute(:applyFill,         :bool)
    define_attribute(:applyBorder,       :bool)
    define_attribute(:applyAlignment,    :bool)
    define_attribute(:applyProtection,   :bool)
    define_child_node(RubyXL::Alignment)
    define_child_node(RubyXL::Protection)
    define_element_name 'xf'

    def ==(other)
      (self.num_fmt_id == other.num_fmt_id) &&
        (self.font_id    == other.font_id) &&
        (self.fill_id    == other.fill_id) &&
        (self.border_id  == other.border_id) &&
        (self.apply_number_format == other.apply_number_format) &&
        (self.apply_font          == other.apply_font) &&
        (self.apply_fill          == other.apply_fill) &&
        (self.apply_border        == other.apply_border) &&
        (self.apply_alignment     == other.apply_alignment) &&
        (self.apply_protection    == other.apply_protection) &&
        (self.xf_id               == other.xf_id)
    end
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_numFmt-1.html
  class NumFmt < OOXMLObject
    define_attribute(:numFmtId,   :int,    :required => true)
    define_attribute(:formatCode, :string, :required => true)
    define_element_name 'numFmt'
  end

end
