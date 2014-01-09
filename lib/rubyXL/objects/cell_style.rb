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
    define_attribute(:horizontal,      :string,
                       :values => %w{general left center right fill justify centerContinuous distributed})
    define_attribute(:vertical,        :string,
                       :values => %w{top center bottom justify distributed})
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

=begin
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
=end
=begin
<xf numFmtId="14" fontId="60" fillId="11" borderId="22" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1">
<alignment horizontal="left"/>
<protection locked="0"/>
</xf>
=end
  end

end
