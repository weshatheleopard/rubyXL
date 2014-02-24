# encoding: UTF-8  <-- magic comment, need this because of sime fancy fonts in the default scheme below. See http://stackoverflow.com/questions/6444826/ruby-utf-8-file-encoding
require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/extensions'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-a_ext-1.html
  class AExtension < OOXMLObject
    define_attribute(:uri, :string)
    define_element_name 'a:ext'
    attr_accessor :raw_xml

    def self.parse(node)
      obj = new
      obj.raw_xml = node.to_xml
      obj
    end

    def write_xml(xml, node_name_override = nil)
      self.raw_xml
    end

  end

  class AExtensionStorageArea < OOXMLObject
    define_child_node(RubyXL::AExtension, :collection => true)
    define_element_name 'a:extLst'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_srgbClr-1.html
  class CT_ScRgbColor < OOXMLObject
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:tint')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:shade')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:comp')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:inv')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gray')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alpha')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:sat')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lum')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:red')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:green')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueMod')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gamma')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:invGamma')
    define_attribute(:r, :int, :required => true)
    define_attribute(:g, :int, :required => true)
    define_attribute(:b, :int, :required => true)
    define_element_name 'a:scrgbClr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_srgbClr-1.html
  class CT_SRgbColor < OOXMLObject
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:tint')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:shade')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:comp')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:inv')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gray')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alpha')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:sat')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lum')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:red')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:green')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueMod')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gamma')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:invGamma')
    define_attribute(:val, :string, :required => true)
    define_element_name 'a:srgbClr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_hslClr-1.html
  class CT_HslColor < OOXMLObject
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:tint')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:shade')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:comp')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:inv')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gray')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alpha')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:sat')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lum')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:red')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:green')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueMod')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gamma')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:invGamma')
    define_attribute(:hue, :int, :required => true)
    define_attribute(:sat, :int, :required => true)
    define_attribute(:lum, :int, :required => true)
    define_element_name 'a:hslClr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_sysClr-1.html
  class CT_SystemColor < OOXMLObject
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:tint')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:shade')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:comp')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:inv')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gray')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alpha')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:sat')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lum')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:red')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:green')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueMod')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gamma')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:invGamma')
    define_attribute(:val,     RubyXL::ST_SystemColorVal, :required => true)
    define_attribute(:lastClr, :string)
    define_element_name 'a:sysClr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_schemeClr-1.html
  class CT_SchemeColor < OOXMLObject
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:tint')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:shade')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:comp')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:inv')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gray')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alpha')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:sat')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lum')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:red')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:green')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueMod')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gamma')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:invGamma')
    define_attribute(:val, RubyXL::ST_SchemeColorVal, :required => true)
    define_element_name 'a:schemeClr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_prstClr-1.html
  class CT_PresetColor < OOXMLObject
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:tint')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:shade')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:comp')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:inv')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gray')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alpha')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:alphaMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:hueMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:sat')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:satMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lum')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:lumMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:red')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:redMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:green')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:greenMod')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blue')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueOff')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:blueMod')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:gamma')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:invGamma')
    define_attribute(:val, RubyXL::ST_PresetColorVal, :required => true)
    define_element_name 'a:prstClr'
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_Color.html
  class CT_Color < OOXMLObject
    define_child_node(RubyXL::CT_ScRgbColor)
    define_child_node(RubyXL::CT_SRgbColor)
    define_child_node(RubyXL::CT_HslColor)
    define_child_node(RubyXL::CT_SystemColor)
    define_child_node(RubyXL::CT_SchemeColor)
    define_child_node(RubyXL::CT_PresetColor)
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_clrScheme-1.html
  class CT_ColorScheme < OOXMLObject
    define_child_node(RubyXL::CT_Color, :node_name => 'a:dk1')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:lt1')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:dk2')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:lt2')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:accent1')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:accent2')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:accent3')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:accent4')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:accent5')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:accent6')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:hlink')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:folHlink')
    define_attribute(:name, :string, :required => true)
    define_element_name 'a:clrScheme'
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_SupplementalFont.html
  class CT_SupplementalFont < OOXMLObject
    define_attribute(:script,   :string, :required => true)
    define_attribute(:typeface, :string, :required => true)
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_TextFont.html
  class CT_TextFont < OOXMLObject
    define_attribute(:typeface,    :string)
    define_attribute(:panose,      :string)
    define_attribute(:pitchFamily, :int, :default => 0)
    define_attribute(:charset,     :int, :default => 1)
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_FontCollection.html
  class CT_FontCollection < OOXMLObject
    define_child_node(RubyXL::CT_TextFont,         :node_name => 'a:latin')
    define_child_node(RubyXL::CT_TextFont,         :node_name => 'a:ea')
    define_child_node(RubyXL::CT_TextFont,         :node_name => 'a:cs')
    define_child_node(RubyXL::CT_SupplementalFont, :node_name => 'a:font', :collection => true)
    define_child_node(RubyXL::AExtensionStorageArea)
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_fontScheme-1.html
  class FontScheme < OOXMLObject
    define_child_node(RubyXL::CT_FontCollection, :node_name => 'a:majorFont')
    define_child_node(RubyXL::CT_FontCollection, :node_name => 'a:minorFont')
    define_child_node(RubyXL::AExtensionStorageArea)
    define_attribute(:name, :string, :required => true)
    define_element_name 'a:fontScheme'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_gs-1.html
  class CT_GradientStop < OOXMLObject
    define_child_node(RubyXL::CT_ScRgbColor)
    define_child_node(RubyXL::CT_SRgbColor)
    define_child_node(RubyXL::CT_HslColor)
    define_child_node(RubyXL::CT_SystemColor)
    define_child_node(RubyXL::CT_SchemeColor)
    define_child_node(RubyXL::CT_PresetColor)
    define_attribute(:pos, :int, :required => true)
    define_element_name 'a:gs'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_gsLst-1.html
  class CT_GradientStopList < OOXMLContainerObject
    define_child_node(RubyXL::CT_GradientStop, :collection => true, :min => 2)
    define_element_name 'a:gsLst'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_lin-1.html
  class CT_LinearShadeProperties < OOXMLObject
    define_attribute(:ang,    :int)
    define_attribute(:scaled, :bool)
    define_element_name 'a:tileRect'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_tileRect-1.html
  class CT_RelativeRect < OOXMLObject
    define_attribute(:l, :int, :default => 0)
    define_attribute(:t, :int, :default => 0)
    define_attribute(:r, :int, :default => 0)
    define_attribute(:b, :int, :default => 0)
    define_element_name 'a:tileRect'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_path-1.html
  class CT_PathShadeProperties < OOXMLObject
    define_child_node(CT_RelativeRect, :node_name => 'a:fillToRect')
    define_attribute(:path, RubyXL::ST_PathShadeType)
    define_element_name 'a:path'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_gradFill-1.html
  class CT_GradientFillProperties < OOXMLObject
    define_child_node(RubyXL::CT_GradientStopList)
    define_child_node(RubyXL::CT_LinearShadeProperties)
    define_child_node(RubyXL::CT_PathShadeProperties)
    define_child_node(RubyXL::CT_RelativeRect)
    define_attribute(:flip,         RubyXL::ST_TileFlipMode)
    define_attribute(:rotWithShape, :bool)
    define_element_name 'a:gradFill'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_pattFill-1.html
  class CT_PatternFillProperties < OOXMLObject
    define_child_node(RubyXL::CT_Color, :node_name => 'a:fgClr')
    define_child_node(RubyXL::CT_Color, :node_name => 'a:bgClr')
    define_attribute(:prst, RubyXL::ST_PresetPatternVal)
    define_element_name 'a:pattFill'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_tile-1.html
  class CT_TileInfoProperties < OOXMLObject
    define_attribute(:tx,    :int)
    define_attribute(:ty,    :int)
    define_attribute(:sx,    :int)
    define_attribute(:sy,    :int)
    define_attribute(:flip,  RubyXL::ST_TileFlipMode)
    define_attribute(:align, RubyXL::ST_RectAlignment)
    define_element_name 'a:tile'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_stretch-1.html
  class CT_StretchInfoProperties < OOXMLObject
    define_child_node(RubyXL::CT_RelativeRect, :node_name => 'a:fillRect')
    define_element_name 'a:stretch'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_blip-1.html
  class CT_Blip < OOXMLObject
#    a:alphaBiLevel    Alpha Bi-Level Effect
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:alphaCeiling')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:alphaFloor')
    define_child_node(RubyXL::CT_Color,     :node_name => 'a:alphaInv')
#    a:alphaMod    Alpha Modulate Effect
#    a:alphaModFix    Alpha Modulate Fixed Effect
#    a:alphaRepl    Alpha Replace Effect
#    a:biLevel    Bi-Level (Black/White) Effect
#    a:blur    Blur Effect
#    a:clrChange    Color Change Effect
    define_child_node(RubyXL::CT_Color,     :node_name => 'a:clrRepl')
#    a:duotone    Duotone Effect
    define_child_node(RubyXL::CT_Color,     :node_name => 'a:fillOverlay')
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:grayscl')
#    a:hsl    Hue Saturation Luminance Effect
#    a:lum    Luminance Effect
#    a:tint    Tint Effect
    define_attribute(:'r:embed', :string)
    define_attribute(:'r:link',  :string)
    define_attribute(:cstate,    RubyXL::ST_BlipCompression)
    define_element_name 'a:blip'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_blipFill-1.html
  class CT_BlipFillProperties < OOXMLObject
    define_child_node(RubyXL::CT_Blip)
    define_child_node(RubyXL::CT_RelativeRect, :node_name => 'a:srcRect')
    define_child_node(RubyXL::CT_TileInfoProperties)
    define_child_node(RubyXL::CT_StretchInfoProperties)
    define_attribute(:dpi,          :int)
    define_attribute(:rotWithShape, :bool)
    define_element_name 'a:blipFill'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_fillStyleLst-1.html
  class CT_FillStyleList < OOXMLObject
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:noFill')
    define_child_node(RubyXL::CT_Color,     :node_name => 'a:solidFill')
    define_child_node(RubyXL::CT_GradientFillProperties)
    define_child_node(RubyXL::CT_BlipFillProperties)
    define_child_node(RubyXL::CT_PatternFillProperties)
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:grpFill')
    define_element_name 'a:fillStyleLst'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_fmtScheme-1.html
  class CT_StyleMatrix < OOXMLObject
    define_child_node(RubyXL::CT_FillStyleList)
#    a:lnStyleLst [1..1]    Line Style List
#    a:effectStyleLst [1..1]    Effect Style List
#    a:bgFillStyleLst [1..1]    Background Fill Style List    define_element_name 'a:fontScheme'
    define_attribute(:name, :string)
    define_element_name 'a:fmtScheme'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_themeElements-1.html
  class ThemeElements < OOXMLObject
    define_child_node(RubyXL::CT_ColorScheme)
    define_child_node(RubyXL::FontScheme)
    define_child_node(RubyXL::CT_StyleMatrix)
    define_child_node(RubyXL::AExtensionStorageArea)
    define_element_name 'a:themeElements'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_off-1.html
  class Offset < OOXMLObject
    define_attribute(:x, :int, :required => true)
    define_attribute(:y, :int, :required => true)
    define_element_name 'a:off'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_ext-2.html
  class Extents < OOXMLObject
    define_attribute(:cx, :int, :required => true)
    define_attribute(:cy, :int, :required => true)
    define_element_name 'a:ext'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_xfrm-4.html
  class CT_Transform2D < OOXMLObject
    define_attribute(:rot,   :int,  :default => 0)
    define_attribute(:flipH, :bool, :default => false)
    define_attribute(:flipV, :bool, :default => false)
    define_child_node(RubyXL::Offset)
    define_child_node(RubyXL::Extents)
    define_element_name 'a:xfrm'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_gd-1.html
  class ShapeGuide < OOXMLObject
    define_attribute(:name, :string, :required => true)
    define_attribute(:fmla, :string, :required => true)
    define_element_name 'a:gd'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_avLst-1.html
  class CT_GeomGuideList < OOXMLContainerObject
    define_child_node(RubyXL::ShapeGuide, :collection => true)
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_rect-1.html
  class ShapeTextRectangle < OOXMLObject
    define_attribute(:l, :int, :required => true)
    define_attribute(:t, :int, :required => true)
    define_attribute(:r, :int, :required => true)
    define_attribute(:b, :int, :required => true)
    define_element_name 'a:rect'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_pos-2.html
  class CT_AdjPoint2D < OOXMLObject
    define_attribute(:x, :int, :required => true)
    define_attribute(:y, :int, :required => true)
    define_element_name 'a:pos'
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_PolarAdjustHandle.html
  class CT_XYAdjustHandle < OOXMLObject
    define_child_node(RubyXL::CT_AdjPoint2D)
    define_attribute(:gdRefX,   :string)
    define_attribute(:minX,     :int)
    define_attribute(:maxX,     :int)
    define_attribute(:gdRefY, :string)
    define_attribute(:minY,   :int)
    define_attribute(:maxY,   :int)
    define_element_name 'a:ahXY'
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_PolarAdjustHandle.html
  class CT_PolarAdjustHandle < OOXMLObject
    define_child_node(RubyXL::CT_AdjPoint2D)
    define_attribute(:gdRefR,   :string)
    define_attribute(:minR,     :int)
    define_attribute(:maxR,     :int)
    define_attribute(:gdRefAng, :string)
    define_attribute(:minAng,   :int)
    define_attribute(:maxAng,   :int)
    define_element_name 'a:ahPolar'
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_AdjustHandleList.html
  class AdjustHandleList < OOXMLObject
    define_child_node(RubyXL::CT_XYAdjustHandle)
    define_child_node(RubyXL::CT_PolarAdjustHandle)
    define_element_name 'a:ahLst'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_cxn-1.html
  class CT_ConnectionSite < OOXMLObject
    define_child_node(RubyXL::CT_AdjPoint2D)
    define_attribute(:ang, :int)
    define_element_name 'a:cxn'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_cxnLst-1.html
  class CT_ConnectionSiteList < OOXMLContainerObject
    define_child_node(RubyXL::CT_ConnectionSite, :collection => true)
    define_element_name 'a:cxnLst'
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_Path2DLineTo.html
  class CT_Path2DTo < OOXMLContainerObject
    define_child_node(RubyXL::CT_AdjPoint2D)
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_arcTo-1.html
  class CT_Path2DArcTo < OOXMLObject
    define_attribute(:wR,    :int, :required => true)
    define_attribute(:hR,    :int, :required => true)
    define_attribute(:stAng, :int, :required => true)
    define_attribute(:swAng, :int, :required => true)
    define_element_name 'a:arcTo'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_quadBezTo-1.html
  class CT_Path2DQuadBezierTo < OOXMLContainerObject
    define_child_node(RubyXL::CT_AdjPoint2D, :collection => true, :node_name => 'a:pt')
    define_element_name 'a:quadBezTo'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_quadBezTo-1.html
  class CT_Path2DCubicBezierTo < OOXMLContainerObject
    define_child_node(RubyXL::CT_AdjPoint2D, :collection => true, :node_name => 'a:pt')
    define_element_name 'a:cubicBezTo'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_path-2.html
  class CT_Path2D < OOXMLObject
    define_child_node(RubyXL::BooleanValue,           :node_name => 'a:close')
    define_child_node(RubyXL::CT_Path2DTo,            :node_name => 'a:moveTo')
    define_child_node(RubyXL::CT_Path2DTo,            :node_name => 'a:lnTo')
    define_child_node(RubyXL::CT_Path2DArcTo,         :node_name => 'a:arcTo')
    define_child_node(RubyXL::CT_Path2DQuadBezierTo)
    define_child_node(RubyXL::CT_Path2DCubicBezierTo)
    define_attribute(:w,           :int,  :default => 0)
    define_attribute(:h,           :int,  :default => 0)
    define_attribute(:fill,        RubyXL::ST_PathFillMode, :default => 'norm')
    define_attribute(:stroke,      :bool, :default => true)
    define_attribute(:extrusionOk, :bool, :default => true)
    define_element_name 'a:cxn'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_pathLst-1.html
  class CT_Path2DList < OOXMLContainerObject
    define_child_node(RubyXL::CT_Path2D, :collection => true)
    define_element_name 'a:pathLst'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_custGeom-1.html
  class CustomGeometry < OOXMLObject
    define_child_node(RubyXL::CT_GeomGuideList, :node_name => 'a:avLst')
    define_child_node(RubyXL::CT_GeomGuideList, :node_name => 'a:gdLst')
    define_child_node(RubyXL::AdjustHandleList)
    define_child_node(RubyXL::CT_ConnectionSiteList)
    define_child_node(RubyXL::ShapeTextRectangle)
    define_child_node(RubyXL::CT_Path2DList)
    define_element_name 'a:custGeom'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_prstGeom-1.html
  class PresetGeometry < OOXMLObject
    define_child_node(RubyXL::CT_GeomGuideList, :node_name => 'a:avLst')
    define_attribute(:prst, RubyXL::ST_ShapeType, :required => true)
    define_element_name 'a:prstGeom'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_spPr-1.html
  class VisualProperties < OOXMLObject
    define_child_node(RubyXL::CT_Transform2D)
    define_child_node(RubyXL::CustomGeometry)
    define_child_node(RubyXL::PresetGeometry)
#        a:noFill    No Fill
#        a:solidFill    Solid Fill
#        a:gradFill    Gradient Fill
#        a:blipFill    Picture Fill
#        a:pattFill    Pattern Fill
#        a:grpFill    Group Fill
#    a:ln [0..1]    
#        a:effectLst    Effect Container
#        a:effectDag    Effect Container
#    a:scene3d [0..1]    3-D Scene
#    a:sp3d [0..1]    3-D Shape Properties
#    a:extLst [0..1]    Extension List
    define_attribute(:bwMode, RubyXL::ST_BlackWhiteMode)

    define_child_node(RubyXL::AExtensionStorageArea)
    define_element_name 'a:spPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_spDef-1.html
  class ShapeDefault < OOXMLObject
    define_child_node(RubyXL::VisualProperties)
#    a:bodyPr [1..1]    BodyProperties
#    a:lstStyle [1..1]    TextListStyles
#    a:style [0..1]    Shape Style
    define_child_node(RubyXL::AExtensionStorageArea)
    define_element_name 'a:spDef'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_objectDefaults-1.html
  class CT_ObjectStyleDefaults < OOXMLObject
    define_child_node(RubyXL::ShapeDefault)
#    a:spDef [0..1]    
#    a:lnDef [0..1]    LineDefault
#    a:txDef [0..1]    TextDefault
    define_child_node(RubyXL::AExtensionStorageArea)
    define_element_name 'a:objectDefaults'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_clrMap-1.html
  class CT_ColorMapping < OOXMLObject
    define_child_node(RubyXL::AExtensionStorageArea)
    define_attribute(:bg1,      RubyXL::ST_ColorSchemeIndex, :required => true)
    define_attribute(:tx1,      RubyXL::ST_ColorSchemeIndex, :required => true)
    define_attribute(:bg2,      RubyXL::ST_ColorSchemeIndex, :required => true)
    define_attribute(:tx2,      RubyXL::ST_ColorSchemeIndex, :required => true)
    define_attribute(:accent1,  RubyXL::ST_ColorSchemeIndex, :required => true)
    define_attribute(:accent2,  RubyXL::ST_ColorSchemeIndex, :required => true)
    define_attribute(:accent3,  RubyXL::ST_ColorSchemeIndex, :required => true)
    define_attribute(:accent4,  RubyXL::ST_ColorSchemeIndex, :required => true)
    define_attribute(:accent5,  RubyXL::ST_ColorSchemeIndex, :required => true)
    define_attribute(:accent6,  RubyXL::ST_ColorSchemeIndex, :required => true)
    define_attribute(:hlink,    RubyXL::ST_ColorSchemeIndex, :required => true)
    define_attribute(:golHlink, RubyXL::ST_ColorSchemeIndex, :required => true)
    define_element_name 'a:clrMap'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_extraClrScheme-1.html
  class CT_ColorSchemeAndMapping < OOXMLObject
    define_child_node(RubyXL::CT_ColorScheme)
    define_child_node(RubyXL::CT_ColorMapping)
    define_element_name 'a:extraClrScheme'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_extraClrSchemeLst-1.html
  class ExtraColorSchemeList < OOXMLContainerObject
    define_child_node(RubyXL::CT_ColorSchemeAndMapping, :collection => true)
    define_element_name 'a:extraClrSchemeLst'
  end
  
  # http://www.schemacentral.com/sc/ooxml/e-a_custClr-1.html
  class CustomColor < OOXMLObject
    define_child_node(RubyXL::CT_ScRgbColor)
    define_child_node(RubyXL::CT_SRgbColor)
    define_child_node(RubyXL::CT_HslColor)
    define_child_node(RubyXL::CT_SystemColor)
    define_child_node(RubyXL::CT_SchemeColor)
    define_child_node(RubyXL::CT_PresetColor)
    define_attribute(:name, :string, :default => '')
    define_element_name 'a:custClr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_custClrLst-1.html
  class CustomColorList < OOXMLContainerObject
    define_child_node(RubyXL::CustomColor, :collection => true)
    define_element_name 'a:custClrLst'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_theme.html
  class Theme < OOXMLTopLevelObject
    define_attribute(:name, :string, :default => '')
    define_child_node(RubyXL::ThemeElements)
    define_child_node(RubyXL::CT_ObjectStyleDefaults)
    define_child_node(RubyXL::ExtraColorSchemeList)
    define_child_node(RubyXL::CustomColorList)
    define_child_node(RubyXL::AExtensionStorageArea)

    define_element_name 'a:theme'

    set_namespaces('xmlns:a' => 'http://schemas.openxmlformats.org/drawingml/2006/main')

    def self.xlsx_path
      File.join('xl', 'theme', 'theme1.xml')
    end

    def self.content_type
      'application/vnd.openxmlformats-officedocument.theme+xml'
    end


    ###### Temporary storage of the theme until I'm done with fully implementing
    ###### all of its intricacies
    attr_accessor :raw_xml

    def self.parse_file(dirpath)
      full_path = File.join(dirpath, xlsx_path)
      return nil unless File.exist?(full_path)

#      test = super
#      puts test.inspect

      obj = self.new
      
      obj.raw_xml = File.open(full_path, 'r').read
      obj
    end

    def write_xml
      raw_xml || # Use fallback theme.
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1>
<a:sysClr val="windowText" lastClr="000000"/>
</a:dk1>
<a:lt1>
<a:sysClr val="window" lastClr="FFFFFF"/>
</a:lt1>
<a:dk2>
<a:srgbClr val="1F497D"/>
</a:dk2>
<a:lt2>
<a:srgbClr val="EEECE1"/>
</a:lt2>
<a:accent1>
<a:srgbClr val="4F81BD"/>
</a:accent1>
<a:accent2>
<a:srgbClr val="C0504D"/>
</a:accent2>
<a:accent3>
<a:srgbClr val="9BBB59"/>
</a:accent3>
<a:accent4>
<a:srgbClr val="8064A2"/>
</a:accent4>
<a:accent5>
<a:srgbClr val="4BACC6"/>
</a:accent5>
<a:accent6>
<a:srgbClr val="F79646"/>
</a:accent6>
<a:hlink>
<a:srgbClr val="0000FF"/>
</a:hlink>
<a:folHlink>
<a:srgbClr val="800080"/>
</a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont>
<a:latin typeface="Cambria"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
<a:font script="Hang" typeface="맑은 고딕"/>
<a:font script="Hans" typeface="宋体"/>
<a:font script="Hant" typeface="新細明體"/>
<a:font script="Arab" typeface="Times New Roman"/>
<a:font script="Hebr" typeface="Times New Roman"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="MoolBoran"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Times New Roman"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
</a:majorFont>
<a:minorFont>
<a:latin typeface="Calibri"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>
<a:font script="Hang" typeface="맑은 고딕"/>
<a:font script="Hans" typeface="宋体"/>
<a:font script="Hant" typeface="新細明體"/>
<a:font script="Arab" typeface="Arial"/>
<a:font script="Hebr" typeface="Arial"/>
<a:font script="Thai" typeface="Tahoma"/>
<a:font script="Ethi" typeface="Nyala"/>
<a:font script="Beng" typeface="Vrinda"/>
<a:font script="Gujr" typeface="Shruti"/>
<a:font script="Khmr" typeface="DaunPenh"/>
<a:font script="Knda" typeface="Tunga"/>
<a:font script="Guru" typeface="Raavi"/>
<a:font script="Cans" typeface="Euphemia"/>
<a:font script="Cher" typeface="Plantagenet Cherokee"/>
<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
<a:font script="Tibt" typeface="Microsoft Himalaya"/>
<a:font script="Thaa" typeface="MV Boli"/>
<a:font script="Deva" typeface="Mangal"/>
<a:font script="Telu" typeface="Gautami"/>
<a:font script="Taml" typeface="Latha"/>
<a:font script="Syrc" typeface="Estrangelo Edessa"/>
<a:font script="Orya" typeface="Kalinga"/>
<a:font script="Mlym" typeface="Kartika"/>
<a:font script="Laoo" typeface="DokChampa"/>
<a:font script="Sinh" typeface="Iskoola Pota"/>
<a:font script="Mong" typeface="Mongolian Baiti"/>
<a:font script="Viet" typeface="Arial"/>
<a:font script="Uigh" typeface="Microsoft Uighur"/>
</a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="50000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="35000">
<a:schemeClr val="phClr">
<a:tint val="37000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:tint val="15000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="1"/>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="100000"/>
<a:shade val="100000"/>
<a:satMod val="130000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:tint val="50000"/>
<a:shade val="100000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:lin ang="16200000" scaled="0"/>
</a:gradFill>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr">
<a:shade val="95000"/>
<a:satMod val="105000"/>
</a:schemeClr>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:prstDash val="solid"/>
</a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="38000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
</a:effectStyle>
<a:effectStyle>
<a:effectLst>
<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
<a:srgbClr val="000000">
<a:alpha val="35000"/>
</a:srgbClr>
</a:outerShdw>
</a:effectLst>
<a:scene3d>
<a:camera prst="orthographicFront">
<a:rot lat="0" lon="0" rev="0"/>
</a:camera>
<a:lightRig rig="threePt" dir="t">
<a:rot lat="0" lon="0" rev="1200000"/>
</a:lightRig>
</a:scene3d>
<a:sp3d>
<a:bevelT w="63500" h="25400"/>
</a:sp3d>
</a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill>
<a:schemeClr val="phClr"/>
</a:solidFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="40000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="40000">
<a:schemeClr val="phClr">
<a:tint val="45000"/>
<a:shade val="99000"/>
<a:satMod val="350000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="20000"/>
<a:satMod val="255000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
</a:path>
</a:gradFill>
<a:gradFill rotWithShape="1">
<a:gsLst>
<a:gs pos="0">
<a:schemeClr val="phClr">
<a:tint val="80000"/>
<a:satMod val="300000"/>
</a:schemeClr>
</a:gs>
<a:gs pos="100000">
<a:schemeClr val="phClr">
<a:shade val="30000"/>
<a:satMod val="200000"/>
</a:schemeClr>
</a:gs>
</a:gsLst>
<a:path path="circle">
<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
</a:path>
</a:gradFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
<a:objectDefaults>
<a:spDef>
<a:spPr/>
<a:bodyPr/>
<a:lstStyle/>
<a:style>
<a:lnRef idx="1">
<a:schemeClr val="accent1"/>
</a:lnRef>
<a:fillRef idx="3">
<a:schemeClr val="accent1"/>
</a:fillRef>
<a:effectRef idx="2">
<a:schemeClr val="accent1"/>
</a:effectRef>
<a:fontRef idx="minor">
<a:schemeClr val="lt1"/>
</a:fontRef>
</a:style>
</a:spDef>
<a:lnDef>
<a:spPr/>
<a:bodyPr/>
<a:lstStyle/>
<a:style>
<a:lnRef idx="2">
<a:schemeClr val="accent1"/>
</a:lnRef>
<a:fillRef idx="0">
<a:schemeClr val="accent1"/>
</a:fillRef>
<a:effectRef idx="1">
<a:schemeClr val="accent1"/>
</a:effectRef>
<a:fontRef idx="minor">
<a:schemeClr val="tx1"/>
</a:fontRef>
</a:style>
</a:lnDef>
</a:objectDefaults>
<a:extraClrSchemeLst/>
</a:theme>'

    end
    ######

  end

end
