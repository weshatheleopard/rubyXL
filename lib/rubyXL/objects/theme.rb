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
#FIXME#
    define_child_node(RubyXL::AExtension, :collection => true)
    define_element_name 'a:extLst'
  end

  class ColorScheme < OOXMLObject
#FIXME#
    define_element_name 'a:clrScheme'
  end

  class FontScheme < OOXMLObject
#FIXME#
    define_element_name 'a:fontScheme'
  end

  class FormatScheme < OOXMLObject
#FIXME#
    define_element_name 'a:fmtScheme'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_themeElements-1.html
  class ThemeElements < OOXMLObject
    define_child_node(RubyXL::ColorScheme)
    define_child_node(RubyXL::FontScheme)
    define_child_node(RubyXL::FormatScheme)
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
  class TwoDTransform < OOXMLObject
    define_attribute(:rot,   :int,  :default => 0)
    define_attribute(:flipH, :bool, :default => false)
    define_attribute(:flipV, :bool, :default => false)
    define_child_node(RubyXL::Offset)
    define_child_node(RubyXL::Extents)
    define_element_name 'a:xfrm'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_custGeom-1.html
  class CustomGeometry < OOXMLObject
#        a:avLst [0..1]    Adjust Value List
#        a:gdLst [0..1]    List of Shape Guides
#        a:ahLst [0..1]    List of Shape Adjust Handles
#        a:cxnLst [0..1]    List of Shape Connection Sites
#        a:rect [0..1]    Shape Text Rectangle
#        a:pathLst [1..1]    List of Shape Paths
    define_element_name 'a:custGeom'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_prstGeom-1.html
  class PresetGeometry < OOXMLObject
#        a:avLst [0..1]    Adjust Value List
    define_attribute(:prst, RubyXL::ST_ShapeType, :required => true)
    define_element_name 'a:prstGeom'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_spPr-1.html
  class VisualProperties < OOXMLObject
    define_child_node(RubyXL::TwoDTransform)
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
  class ObjectDefaults < OOXMLObject
    define_child_node(RubyXL::ShapeDefault)
#    a:spDef [0..1]    
#    a:lnDef [0..1]    LineDefault
#    a:txDef [0..1]    TextDefault
    define_child_node(RubyXL::AExtensionStorageArea)
    define_element_name 'a:objectDefaults'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_extraClrScheme-1.html
  class ExtraColorScheme < OOXMLObject
#    a:clrScheme [1..1]    ColorScheme
#    a:clrMap [0..1]    ColorMap
    define_element_name 'a:extraClrScheme'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_extraClrSchemeLst-1.html
  class ExtraColorSchemeList < OOXMLContainerObject
    define_child_node(RubyXL::ExtraColorScheme, :collection => true)
    define_element_name 'a:extraClrSchemeLst'
  end
  
  # http://www.schemacentral.com/sc/ooxml/e-a_custClr-1.html
  class CustomColor < OOXMLObject
#    a:scrgbClr    RGB Color Model - Percentage Variant
#    a:srgbClr    RGB Color Model - Hex Variant
#    a:hslClr    Hue, Saturation, Luminance Color Model
#    a:sysClr    System Color
#    a:schemeClr    Scheme Color
#    a:prstClr    Preset Color
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
    define_child_node(RubyXL::ObjectDefaults)
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
