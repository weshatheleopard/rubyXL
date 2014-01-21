module RubyXL

  # Eventually, the entire code for Worksheet will be moved here. One small step at a time!

  # http://www.schemacentral.com/sc/ooxml/e-ssml_legacyDrawing-1.html
  class LegacyDrawing < OOXMLObject
    define_attribute(:'r:id',            :string, :required => true)
    define_element_name 'legacyDrawing'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetPr-3.html
  class WorksheetProperties < OOXMLObject
    define_attribute(:syncHorizontal,                    :bool, :default => false)
    define_attribute(:syncVertical,                      :bool, :default => false)
    define_attribute(:syncRef,                           :ref)
    define_attribute(:transitionEvaluation,              :bool, :default => false)
    define_attribute(:transitionEntry,                   :bool, :default => false)
    define_attribute(:published,                         :bool, :default => true)
    define_attribute(:codeName,                          :string)
    define_attribute(:filterMode,                        :bool, :default => false)
    define_attribute(:enableFormatConditionsCalculation, :bool, :default => true)
#    define_child_node(RubyXL::TabColor)
#    define_child_node(RubyXL::OutlineProperties)
#    define_child_node(RubyXL::PageSetupProperties)
    define_element_name 'sheetPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_dimension-3.html
  class WorksheetDimensions < OOXMLObject
    define_attribute(:ref, :ref)
    define_element_name 'dimension'
  end

  class WorksheetFormatProperties < OOXMLObject
    define_attribute(:baseColWidth,     :int,  :default => 8)
    define_attribute(:defaultColWidth,  :float)
    define_attribute(:defaultRowHeight, :float)
    define_attribute(:customHeight,     :bool, :default => false)
    define_attribute(:zeroHeight,       :bool, :default => false)
    define_attribute(:thickTop,         :bool, :default => false)
    define_attribute(:thickBottom,      :bool, :default => false)
    define_attribute(:outlineLevelRow,  :int,  :default => 0)
    define_attribute(:outlineLevelCol,  :int,  :default => 0)
    define_element_name 'sheetFormatPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_pageMargins-1.html
  class PageMargins < OOXMLObject
    define_attribute(:left,   :float, :required => true)
    define_attribute(:right,  :float, :required => true)
    define_attribute(:top,    :float, :required => true)
    define_attribute(:bottom, :float, :required => true)
    define_attribute(:header, :float, :required => true)
    define_attribute(:footer, :float, :required => true)
    define_element_name 'pageMargins'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_pageSetup-1.html
  class PageSetup < OOXMLObject
    define_attribute(:paperSize,          :int, :default => 1)
    define_attribute(:scale,              :int, :default => 100)
    define_attribute(:firstPageNumber,    :int, :default => 1)
    define_attribute(:fitToWidth,         :int, :default => 1)
    define_attribute(:fitToHeight,        :int, :default => 1)
    define_attribute(:pageOrder,          :string, :default => 'downThenOver',
                       :values => %w{ downThenOver overThenDown })
    define_attribute(:orientation,        :string, :default => 'default',
                       :values => %w{ default portrait landscape })
    define_attribute(:usePrinterDefaults, :bool, :default => true)
    define_attribute(:blackAndWhite,      :bool, :default => false)
    define_attribute(:draft,              :bool, :default => false)
    define_attribute(:cellComments,       :string, :default => 'none',
                       :values => %w{ none asDisplayed atEnd })
    define_attribute(:useFirstPageNumber, :bool, :default => false)
    define_attribute(:errors,             :string,    :default => 'displayed',
                       :values => %w{ displayed blank dash NA })
    define_attribute(:horizontalDpi,      :int,  :default => 600)
    define_attribute(:verticalDpi,        :int,  :default => 600)
    define_attribute(:copies,             :int,  :default => 1)

    define_attribute(:'r:id',            :string)
    define_element_name 'pageSetup'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_drawing-1.html
  class Drawing < OOXMLObject
    define_attribute(:'r:id',            :string)
    define_element_name 'drawing'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_mergeCell-1.html
  class MergedCell < OOXMLObject
    define_attribute(:ref, :ref)
    define_element_name 'mergeCell'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_mergeCells-1.html
  class MergedCells < OOXMLObject
    define_child_node(RubyXL::MergedCell, :collection => :with_count)
    define_element_name 'mergeCells'
  end

  class Worksheet < OOXMLObject
    define_child_node(RubyXL::WorksheetProperties)
    define_child_node(RubyXL::WorksheetDimensions)
    define_child_node(RubyXL::SheetViews)
    define_child_node(RubyXL::WorksheetFormatProperties)
    define_child_node(RubyXL::ColumnRanges)
    define_child_node(RubyXL::SheetData)
#    ssml:sheetCalcPr [0..1]    Sheet Calculation Properties
#    ssml:sheetProtection [0..1]    Sheet Protection
#    ssml:protectedRanges [0..1]    Protected Ranges
#    ssml:scenarios [0..1]    Scenarios
#    ssml:autoFilter [0..1]    AutoFilter
#    ssml:sortState [0..1]    Sort State
#    ssml:dataConsolidate [0..1]    Data Consolidate
#    ssml:customSheetViews [0..1]    Custom Sheet Views
    define_child_node(RubyXL::MergedCells)
#    ssml:phoneticPr [0..1]    Phonetic Properties
#    ssml:conditionalFormatting [0..*]    Conditional Formatting
    define_child_node(RubyXL::DataValidations)
#    ssml:hyperlinks [0..1]    Hyperlinks
#    ssml:printOptions [0..1]    Print Options
    define_child_node(RubyXL::PageMargins)
    define_child_node(RubyXL::PageSetup)
#    ssml:headerFooter [0..1]    Header Footer Settings
#    ssml:rowBreaks [0..1]    Horizontal Page Breaks
#    ssml:colBreaks [0..1]    Vertical Page Breaks
#    ssml:customProperties [0..1]    Custom Properties
#    ssml:cellWatches [0..1]    Cell Watch Items
#    ssml:ignoredErrors [0..1]    Ignored Errors
#    ssml:smartTags [0..1]    Smart Tags
    define_child_node(RubyXL::Drawing)
#    ssml:legacyDrawing [0..1]    Legacy Drawing
#    ssml:legacyDrawingHF [0..1]    Legacy Drawing Header Footer
#    ssml:picture [0..1]    Background Image
#    ssml:oleObjects [0..1]    OLE Objects
#    ssml:controls [0..1]    Embedded Controls
#    ssml:webPublishItems [0..1]    Web Publishing Items
#    ssml:tableParts [0..1]    Table Parts
#    ssml:extLst [0..1]    Future Feature Storage Area
    define_element_name 'worksheet'

    include LegacyWorksheet
  end

end
