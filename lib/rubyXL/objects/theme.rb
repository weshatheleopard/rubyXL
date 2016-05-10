# encoding: UTF-8  <-- magic comment, need this because of sime fancy fonts in the default scheme below. See http://stackoverflow.com/questions/6444826/ruby-utf-8-file-encoding
require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/extensions'
require 'rubyXL/objects/complex_types'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-a_fontScheme-1.html
  class FontScheme < OOXMLObject
    # -- Sequence [1..1]
    define_child_node(RubyXL::CT_FontCollection, :node_name => 'a:majorFont')
    define_child_node(RubyXL::CT_FontCollection, :node_name => 'a:minorFont')
    define_child_node(RubyXL::AExtensionStorageArea)
    # --
    define_attribute(:name, :string, :required => true)
    define_element_name 'a:fontScheme'
  end


  # http://www.schemacentral.com/sc/ooxml/e-a_themeElements-1.html
  class ThemeElements < OOXMLObject
    define_child_node(RubyXL::CT_ColorScheme)
    define_child_node(RubyXL::FontScheme)
    define_child_node(RubyXL::CT_StyleMatrix)
    define_child_node(RubyXL::AExtensionStorageArea)
    define_element_name 'a:themeElements'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_rect-1.html
  class ShapeTextRectangle < OOXMLObject
    define_attribute(:l, :int, :required => true)
    define_attribute(:t, :int, :required => true)
    define_attribute(:r, :int, :required => true)
    define_attribute(:b, :int, :required => true)
    define_element_name 'a:rect'
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
    define_child_node(RubyXL::CT_ConnectionSite, :collection => [0..-1])
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
    define_child_node(RubyXL::CT_AdjPoint2D, :collection => [2..2], :node_name => 'a:pt')
    define_element_name 'a:quadBezTo'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_quadBezTo-1.html
  class CT_Path2DCubicBezierTo < OOXMLContainerObject
    define_child_node(RubyXL::CT_AdjPoint2D, :collection => [2..2], :node_name => 'a:pt')
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
    define_element_name 'a:path'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_pathLst-1.html
  class CT_Path2DList < OOXMLContainerObject
    define_child_node(RubyXL::CT_Path2D, :collection => [0..-1])
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
    # -- Choice [0..1] (EG_Geometry)
    define_child_node(RubyXL::CustomGeometry)
    define_child_node(RubyXL::PresetGeometry)
    # -- Choice [0..1] (EG_FillProperties)
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:noFill')
    define_child_node(RubyXL::CT_Color,     :node_name => 'a:solidFill')
    define_child_node(RubyXL::CT_GradientFillProperties)
    define_child_node(RubyXL::CT_BlipFillProperties)
    define_child_node(RubyXL::CT_PatternFillProperties)
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:grpFill')
    # --
    define_child_node(RubyXL::CT_LineProperties)
    # -- Choice [0..1] (EG_EffectProperties)
    define_child_node(RubyXL::CT_EffectList)
    define_child_node(RubyXL::CT_EffectContainer, :node_name => 'a:effectDag')
    # --
    define_child_node(RubyXL::CT_Scene3D)
    define_child_node(RubyXL::CT_Shape3D)
    define_child_node(RubyXL::AExtensionStorageArea)
    define_attribute(:bwMode, RubyXL::ST_BlackWhiteMode)
    define_element_name 'a:spPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_prstTxWarp-2.html
  class CT_PresetTextShape < OOXMLObject
    define_child_node(RubyXL::CT_GeomGuideList, :node_name => 'a:avLst')
    define_attribute(:prst, RubyXL::ST_TextShapeType)
    define_element_name 'a:prstTxWarp'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_normAutofit-1.html
  class CT_TextNormalAutofit < OOXMLObject
    define_attribute(:fontScale,      :int, :default => 100000)
    define_attribute(:lnSpcReduction, :int, :default => 0)
    define_element_name 'a:normAutofit'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_flatTx-1.html
  class CT_FlatText < OOXMLObject
    define_attribute(:z, :int, :default => 0)
    define_element_name 'a:flatTx'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_bodyPr-1.html
  class BodyProperties < OOXMLObject
    define_child_node(RubyXL::CT_PresetTextShape)
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:noAutofit')
    define_child_node(RubyXL::CT_TextNormalAutofit)
    define_child_node(RubyXL::BooleanValue, :node_name => 'a:spAutoFit')
    define_child_node(RubyXL::CT_Scene3D)
    define_child_node(RubyXL::CT_Shape3D)
    define_child_node(RubyXL::CT_FlatText)
    define_child_node(RubyXL::AExtensionStorageArea)
    define_attribute(:rot, :int)
    define_attribute(:spcFirstLastPara, :bool)
    define_attribute(:vertOverflow,     RubyXL::ST_TextVertOverflowType)
    define_attribute(:horzOverflow,     RubyXL::ST_TextHorzOverflowType)
    define_attribute(:vert,             RubyXL::ST_TextVerticalType)
    define_attribute(:wrap,             RubyXL::ST_TextWrappingType)
    define_attribute(:lIns,             :int)
    define_attribute(:tIns,             :int)
    define_attribute(:rIns,             :int)
    define_attribute(:bIns,             :int)
    define_attribute(:numCol,           :int)
    define_attribute(:spcCol,           :int)
    define_attribute(:rtlCol,           :bool)
    define_attribute(:fromWordArt,      :bool)
    define_attribute(:anchor,           RubyXL::ST_TextAnchoringType)
    define_attribute(:anchorCtr,        :bool)
    define_attribute(:forceAA,          :bool)
    define_attribute(:upright,          :bool, :default => false)
    define_attribute(:compatLnSpc,      :bool)
    define_element_name 'a:bodyPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_tab-1.html
  class CT_TextTabStop < OOXMLObject
    define_attribute(:pos,  :int)
    define_attribute(:algn, RubyXL::ST_TextTabAlignType)
    define_element_name 'a:tabLst'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_tabLst-1.html
  class CT_TextTabStopList < OOXMLContainerObject
    define_child_node(RubyXL::CT_TextTabStop, :collection => [0..32])
    define_element_name 'a:tabLst'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_buAutoNum-1.html
  class CT_TextAutonumberBullet < OOXMLObject
    define_attribute(:type, RubyXL::ST_TextAutonumberScheme)
    define_attribute(:startAt, :int)
    define_element_name 'a:buAutoNum'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_buChar-1.html
  class CT_TextCharBullet < OOXMLObject
    define_attribute(:char, :string, :required => true)
    define_element_name 'a:buChar'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_buBlip-1.html
  class CT_TextBlipBullet < OOXMLObject
    define_child_node(RubyXL::CT_Blip)
    define_element_name 'a:buBlip'
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_TextSpacing.html
  class CT_TextSpacing < OOXMLObject
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:spcPct')
    define_child_node(RubyXL::IntegerValue, :node_name => 'a:spcPts')
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_snd-1.html
  class CT_EmbeddedWAVAudioFile < OOXMLObject
    define_attribute(:'r:embed', :string)
    define_attribute(:name,      :string, :default => '')
    define_attribute(:builtIn,   :bool,   :default => false)
    define_element_name 'a:snd'
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_Hyperlink.html
  class CT_Hyperlink < OOXMLObject
    define_child_node(RubyXL::CT_EmbeddedWAVAudioFile)
    define_child_node(RubyXL::AExtensionStorageArea)
    define_attribute(:'r:id',         :string)
    define_attribute(:invalidUrl,     :string, :default => '')
    define_attribute(:action,         :string, :default => '')
    define_attribute(:tgtFrame,       :string, :default => '')
    define_attribute(:tooltip,        :string, :default => '')
    define_attribute(:history,        :bool,   :default => true)
    define_attribute(:highlightClick, :bool,   :default => false)
    define_attribute(:endSnd,         :bool,   :default => false)
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_defRPr-1.html
  class CT_TextCharacterProperties < OOXMLObject
    define_child_node(RubyXL::CT_LineProperties)
    # -- EG_FillProperties
    define_child_node(RubyXL::BooleanValue,      :node_name => 'a:noFill')
    define_child_node(RubyXL::CT_Color,          :node_name => 'a:solidFill')
    define_child_node(RubyXL::CT_GradientFillProperties)
    define_child_node(RubyXL::CT_BlipFillProperties)
    define_child_node(RubyXL::CT_PatternFillProperties)
    define_child_node(RubyXL::BooleanValue,      :node_name => 'a:grpFill')
    # -- EG_EffectProperties
    define_child_node(RubyXL::CT_EffectList)
    define_child_node(RubyXL::CT_EffectContainer, :node_name => 'a:effectDag')
    # --
    define_child_node(RubyXL::CT_Color,          :node_name => 'a:highlight')
    # -- EG_TextUnderlineLine
    define_child_node(RubyXL::BooleanValue,      :node_name => 'a:uLnTx')
    define_child_node(RubyXL::CT_LineProperties, :node_name => 'a:uLn')
    # -- EG_TextUnderlineFill
    define_child_node(RubyXL::BooleanValue,      :node_name => 'a:uFillTx')
    define_child_node(RubyXL::CT_FillStyleList,  :node_name => 'a:uFill')
    define_child_node(RubyXL::CT_TextFont,       :node_name => 'a:latin')
    define_child_node(RubyXL::CT_TextFont,       :node_name => 'a:ea')
    define_child_node(RubyXL::CT_TextFont,       :node_name => 'a:cs')
    define_child_node(RubyXL::CT_TextFont,       :node_name => 'a:sym')
    define_child_node(RubyXL::CT_Hyperlink,      :node_name => 'a:hlinkClick')
    define_child_node(RubyXL::CT_Hyperlink,      :node_name => 'a:hlinkMouseOver')
    define_child_node(RubyXL::AExtensionStorageArea)
    define_attribute(:kumimoji,   :bool)
    define_attribute(:lang,       :string)
    define_attribute(:altLang,    :string)
    define_attribute(:sz,         :int)
    define_attribute(:b,          :bool)
    define_attribute(:i,          :bool)
    define_attribute(:u,          RubyXL::ST_TextUnderlineType)
    define_attribute(:strike,     RubyXL::ST_TextStrikeType)
    define_attribute(:kern,       :int)
    define_attribute(:cap,        RubyXL::ST_TextCapsType)
    define_attribute(:spc,        :int)
    define_attribute(:normalizeH, :bool)
    define_attribute(:baseline,   :int)
    define_attribute(:noProof,    :bool)
    define_attribute(:dirty,      :bool, :default => true)
    define_attribute(:err,        :bool, :default => false)
    define_attribute(:smtClean,   :bool, :default => true)
    define_attribute(:smtId,      :int,  :default => 0)
    define_attribute(:bmk,        :string)
    define_element_name 'a:defRPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_defPPr-1.html
  class CT_TextParagraphProperties < OOXMLObject
    define_child_node(RubyXL::CT_TextSpacing, :node_name => 'a:lnSpc')
    define_child_node(RubyXL::CT_TextSpacing, :node_name => 'a:spcBef')
    define_child_node(RubyXL::CT_TextSpacing, :node_name => 'a:spcAft')
    define_child_node(RubyXL::BooleanValue,   :node_name => 'a:buClrTx')
    define_child_node(RubyXL::CT_Color,       :node_name => 'a:buClr')
    define_child_node(RubyXL::BooleanValue,   :node_name => 'a:buSzTx')
    define_child_node(RubyXL::IntegerValue,   :node_name => 'a:buSzPct')
    define_child_node(RubyXL::IntegerValue,   :node_name => 'a:buSzPts')
    define_child_node(RubyXL::BooleanValue,   :node_name => 'a:buFontTx')
    define_child_node(RubyXL::CT_TextFont,    :node_name => 'a:buFont')
    define_child_node(RubyXL::BooleanValue,   :node_name => 'a:buNone')
    define_child_node(RubyXL::CT_TextAutonumberBullet)
    define_child_node(RubyXL::CT_TextCharBullet)
    define_child_node(RubyXL::CT_TextBlipBullet)
    define_child_node(RubyXL::CT_TextTabStop)
    define_child_node(RubyXL::CT_TextCharacterProperties)
    define_child_node(RubyXL::AExtensionStorageArea)
    define_attribute(:marL,         :int)
    define_attribute(:marR,         :int)
    define_attribute(:lvl,          :int)
    define_attribute(:indent,       :int)
    define_attribute(:algn,         RubyXL::ST_TextAlignType)
    define_attribute(:defTabSz,     :int)
    define_attribute(:rtl,          :bool)
    define_attribute(:eaLnBrk,      :bool)
    define_attribute(:fontAlgn,     RubyXL::ST_TextFontAlignType)
    define_attribute(:latinLnBrk,   :bool)
    define_attribute(:hangingPunct, :bool)
    define_element_name 'a:defPPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_lstStyle-1.html
  class CT_TextListStyle < OOXMLObject
    define_child_node(RubyXL::CT_TextParagraphProperties, :node_name => 'a:defPPr')
    define_child_node(RubyXL::CT_TextParagraphProperties, :node_name => 'a:lvl1pPr')
    define_child_node(RubyXL::CT_TextParagraphProperties, :node_name => 'a:lvl2pPr')
    define_child_node(RubyXL::CT_TextParagraphProperties, :node_name => 'a:lvl3pPr')
    define_child_node(RubyXL::CT_TextParagraphProperties, :node_name => 'a:lvl4pPr')
    define_child_node(RubyXL::CT_TextParagraphProperties, :node_name => 'a:lvl5pPr')
    define_child_node(RubyXL::CT_TextParagraphProperties, :node_name => 'a:lvl6pPr')
    define_child_node(RubyXL::CT_TextParagraphProperties, :node_name => 'a:lvl7pPr')
    define_child_node(RubyXL::CT_TextParagraphProperties, :node_name => 'a:lvl8pPr')
    define_child_node(RubyXL::CT_TextParagraphProperties, :node_name => 'a:lvl9pPr')
    define_child_node(RubyXL::AExtensionStorageArea)
    define_element_name 'a:lstStyle'
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_StyleMatrixReference.html
  class CT_StyleMatrixReference < OOXMLObject
    define_child_node(RubyXL::CT_ScRgbColor)
    define_child_node(RubyXL::CT_SRgbColor)
    define_child_node(RubyXL::CT_HslColor)
    define_child_node(RubyXL::CT_SystemColor)
    define_child_node(RubyXL::CT_SchemeColor)
    define_child_node(RubyXL::CT_PresetColor)
    define_attribute(:idx, :int, :required => true)
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_FontReference.html
  class CT_FontReference < OOXMLObject
    define_child_node(RubyXL::CT_ScRgbColor)
    define_child_node(RubyXL::CT_SRgbColor)
    define_child_node(RubyXL::CT_HslColor)
    define_child_node(RubyXL::CT_SystemColor)
    define_child_node(RubyXL::CT_SchemeColor)
    define_child_node(RubyXL::CT_PresetColor)
    define_attribute(:idx, RubyXL::ST_FontCollectionIndex, :required => true)
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_ShapeStyle.html
  class CT_ShapeStyle < OOXMLObject
    define_child_node(RubyXL::CT_StyleMatrixReference, :node_name => 'a:lnRef')
    define_child_node(RubyXL::CT_StyleMatrixReference, :node_name => 'a:fillRef')
    define_child_node(RubyXL::CT_StyleMatrixReference, :node_name => 'a:effectRef')
    define_child_node(RubyXL::CT_FontReference,        :node_name => 'a:fontRef')
    define_element_name 'a:style'
  end

  # http://www.schemacentral.com/sc/ooxml/t-a_CT_DefaultShapeDefinition.html
  class CT_DefaultShapeDefinition < OOXMLObject
    define_child_node(RubyXL::VisualProperties)
    define_child_node(RubyXL::BodyProperties)
    define_child_node(RubyXL::CT_TextListStyle)
    define_child_node(RubyXL::CT_ShapeStyle)
    define_child_node(RubyXL::AExtensionStorageArea)
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_objectDefaults-1.html
  class CT_ObjectStyleDefaults < OOXMLObject
    define_child_node(RubyXL::CT_DefaultShapeDefinition, :node_name => 'a:spDef')
    define_child_node(RubyXL::CT_DefaultShapeDefinition, :node_name => 'a:lnDef')
    define_child_node(RubyXL::CT_DefaultShapeDefinition, :node_name => 'a:txDef')
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
    define_child_node(RubyXL::CT_ColorSchemeAndMapping, :collection => [0..-1])
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
    define_child_node(RubyXL::CustomColor, :collection => [0..-1])
    define_element_name 'a:custClrLst'
  end

  # http://www.schemacentral.com/sc/ooxml/e-a_theme.html
  class Theme < OOXMLTopLevelObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.theme+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'

    define_attribute(:name, :string, :default => '')
    define_child_node(RubyXL::ThemeElements)
    define_child_node(RubyXL::CT_ObjectStyleDefaults)
    define_child_node(RubyXL::ExtraColorSchemeList)
    define_child_node(RubyXL::CustomColorList)
    define_child_node(RubyXL::AExtensionStorageArea)

    define_element_name 'a:theme'

    set_namespaces('http://schemas.openxmlformats.org/drawingml/2006/main' => 'a')

    def xlsx_path
      ROOT.join('xl', 'theme', 'theme1.xml')
    end

    def self.default
      default_theme = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
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
      self.parse(default_theme)
    end

  end

end
