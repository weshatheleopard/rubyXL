module RubyXL

  # Eventually, the entire code for Workbook will be moved here.


  # http://www.schemacentral.com/sc/ooxml/e-ssml_fileVersion-1.html
  class FileVersion < OOXMLObject
    define_attribute(:appName,      :string)
    define_attribute(:lastEdited,   :string)
    define_attribute(:lowestEdited, :string)
    define_attribute(:rupBuild,     :string)
    define_attribute(:codeName,     :string)
    define_element_name 'fileVersion'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_fileSharing-1.html
  class FileSharing < OOXMLObject
    define_attribute(:readOnlyRecommended, :bool, :default => false)
    define_attribute(:userName,            :string)
    define_attribute(:reservationPassword, :string)
    define_element_name 'fileSharing'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_workbookPr-1.html
  class WorkbookProperties < OOXMLObject
    define_attribute(:date1904,                   :bool,   :default => false)
    define_attribute(:showObjects,                :string, :default => 'all', :values =>
                       %w{ all placeholders none } )
    define_attribute(:showBorderUnselectedTables, :bool,   :default => true)
    define_attribute(:filterPrivacy,              :bool,   :default => false)
    define_attribute(:promptedSolutions,          :bool,   :default => false)
    define_attribute(:showInkAnnotation,          :bool,   :default => true)
    define_attribute(:backupFile,                 :bool,   :default => false)
    define_attribute(:saveExternalLinkValues,     :bool,   :default => true)
    define_attribute(:updateLinks,                :string, :default => 'userSet', :values =>
                       %w{ userSet never always } )
    define_attribute(:hidePivotFieldList,         :bool,   :default => false)
    define_attribute(:showPivotChartFilter,       :bool,   :default => false)
    define_attribute(:allowRefreshQuery,          :bool,   :default => false)
    define_attribute(:publishItems,               :bool,   :default => false)
    define_attribute(:checkCompatibility,         :bool,   :default => false)
    define_attribute(:autoCompressPictures,       :bool,   :default => true)
    define_attribute(:refreshAllConnections,      :bool,   :default => false)
    define_attribute(:defaultThemeVersion,        :int)
    define_attribute(:codeName,                   :string)
    define_element_name 'workbookPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_workbookProtection-1.html
  class WorkbookProtection < OOXMLObject
    define_attribute(:workbookPassword,  :string)
    define_attribute(:revisionsPassword, :string)
    define_attribute(:lockStructure,     :bool,   :default => false)
    define_attribute(:lockWindows,       :bool,   :default => false)
    define_attribute(:lockRevision,      :bool,   :default => false)
    define_element_name 'workbookProtection'
  end


  class WorkbookView < OOXMLObject
    define_attribute(:visibility,             :string, :default => 'visible', :values =>
                       %w{ visible hidden veryHidden } )
    define_attribute(:minimized,              :bool,   :default => false)
    define_attribute(:showHorizontalScroll,   :bool,   :default => true)
    define_attribute(:showVerticalScroll,     :bool,   :default => true)
    define_attribute(:showSheetTabs,          :bool,   :default => true)
    define_attribute(:xWindow,                :int)
    define_attribute(:yWindow,                :int)
    define_attribute(:windowWidth,            :int)
    define_attribute(:windowHeight,           :int)
    define_attribute(:tabRatio,               :int,    :default => 600)
    define_attribute(:firstSheet,             :int,    :default => 0)
    define_attribute(:activeTab,              :int,    :default => 0)
    define_attribute(:autoFilterDateGrouping, :bool,   :default => true)
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'workbookView'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_bookViews-1.html
  class WorkbookViews < OOXMLObject
    define_child_node(RubyXL::WorkbookView, :collection => true)
    define_element_name 'bookViews'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheet-1.html
  class Sheet < OOXMLObject
    define_attribute(:name,            :string, :required => true)
    define_attribute(:sheetId,         :int, :required => true)
    define_attribute(:state,             :string, :default => 'visible', :values =>
                       %w{ visible hidden veryHidden } )
    define_attribute(:'r:id',            :string, :required => true)
    define_element_name 'sheet'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheets-1.html
  class Sheets < OOXMLObject
    define_child_node(RubyXL::Sheet, :collection => true)
    define_element_name 'sheets'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_definedName-2.html
  class DefinedName < OOXMLObject
    define_attribute(:name,              :string, :required => true)
    define_attribute(:comment,           :string)
    define_attribute(:customMenu,        :string)
    define_attribute(:description,       :string)
    define_attribute(:help,              :string)
    define_attribute(:description,       :string)
    define_attribute(:localSheetId,      :string)

    define_attribute(:hidden,            :bool, :default => false)
    define_attribute(:function,          :bool, :default => false)
    define_attribute(:vbProcedure,       :bool, :default => false)
    define_attribute(:xlm,               :bool, :default => false)

    define_attribute(:functionGroupId,   :int)
    define_attribute(:shortcutKey,       :string)
    define_attribute(:publishToServer,   :bool, :default => false)
    define_attribute(:workbookParameter, :bool, :default => false)

    define_attribute(:_,                 :string, :accessor => :reference)
    define_element_name 'definedName'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_definedName-2.html
  class DefinedNames < OOXMLObject
    define_child_node(RubyXL::DefinedName, :collection => true, :accessor => :defined_names)
    define_element_name 'definedNames'
  end

  class CalculationProperties < OOXMLObject
    define_attribute(:calcId,   :int)
    define_attribute(:calcMode,              :string, :default => 'auto', :values =>
                       %w{ manual auto autoNoTable } )
    define_attribute(:fullCalcOnLoad,        :bool,   :default => false)
    define_attribute(:refMode,               :string, :default => 'A1', :values =>
                       %w{ A1 R1C1 } )
    define_attribute(:iterate,               :bool,   :default => false)
    define_attribute(:iterateCount,          :int,    :default => 100)
    define_attribute(:iterateDelta,          :float,  :default => 0.001)
    define_attribute(:fullPrecision,         :bool,   :default => true)
    define_attribute(:calcCompleted,         :bool,   :default => true)
    define_attribute(:calcOnSave,            :bool,   :default => true)
    define_attribute(:concurrentCalc,        :bool,   :default => true)
    define_attribute(:concurrentManualCount, :int)
    define_attribute(:forceFullCalc,         :bool)
    define_element_name 'calcPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_workbook.html
  class Workbook < OOXMLObject
    define_child_node(RubyXL::FileVersion)
    define_child_node(RubyXL::FileSharing)
    define_child_node(RubyXL::WorkbookProperties, :accessor => :workbook_properties)
    define_child_node(RubyXL::WorkbookProtection)
    define_child_node(RubyXL::WorkbookViews)
    define_child_node(RubyXL::Sheets)
#    ssml:functionGroups [0..1]    Function Groups
#    ssml:externalReferences [0..1]    External References
    define_child_node(RubyXL::DefinedNames, :accessor => :defined_name_container)
    define_child_node(RubyXL::CalculationProperties)
#    ssml:oleSize [0..1]    OLE Size
#    ssml:customWorkbookViews [0..1]    Custom Workbook Views
#    ssml:pivotCaches [0..1]    PivotCaches
#    ssml:smartTagPr [0..1]    Smart Tag Properties
#    ssml:smartTagTypes [0..1]    Smart Tag Types
#    ssml:webPublishing [0..1]    Web Publishing Properties
#    ssml:fileRecoveryPr [0..*]    File Recovery Properties
#    ssml:webPublishObjects [0..1]    Web Publish Objects
    define_child_node(RubyXL::ExtensionStorageArea)
    define_child_node(RubyXL::AlternateContent)

    define_element_name 'workbook'
    set_namespaces('xmlns'     => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                   'xmlns:r'   => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                   'xmlns:mc'  => 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                   'xmlns:x15' => 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main')

    include LegacyWorkbook

    def date1904
      workbook_properties && workbook_properties.date1904
    end

    def date1904=(v)
      self.workbook_properties ||= RubyXL::WorkbookProperties.new
      workbook_properties.date1904 = v
    end


  end


end