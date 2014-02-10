require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/simple_types'
require 'rubyXL/objects/extensions'

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
    define_attribute(:showObjects,                :string, :default => 'all',
                       :values => RubyXL::ST_Objects)
    define_attribute(:showBorderUnselectedTables, :bool,   :default => true)
    define_attribute(:filterPrivacy,              :bool,   :default => false)
    define_attribute(:promptedSolutions,          :bool,   :default => false)
    define_attribute(:showInkAnnotation,          :bool,   :default => true)
    define_attribute(:backupFile,                 :bool,   :default => false)
    define_attribute(:saveExternalLinkValues,     :bool,   :default => true)
    define_attribute(:updateLinks,                :string, :default => 'userSet',
                       :values => RubyXL::ST_UpdateLinks)
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

  # http://www.schemacentral.com/sc/ooxml/e-ssml_workbookView-1.html
  class WorkbookView < OOXMLObject
    define_attribute(:visibility,             :string, :default => 'visible',
                      :values => RubyXL::ST_Visibility)
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
    define_attribute(:sheetId,         :int,    :required => true)
    define_attribute(:state,           :string, :default => 'visible',
                       :values => RubyXL::ST_Visibility)
    define_attribute(:'r:id',          :string, :required => true)
    define_element_name 'sheet'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheets-1.html
  class Sheets < OOXMLObject
    define_child_node(RubyXL::Sheet, :collection => true, :accessor => :sheets)
    define_element_name 'sheets'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_externalReference-1.html
  class ExternalReference < OOXMLObject
    define_attribute(:'r:id', :string, :required => true)
    define_element_name 'externalReference'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_externalReferences-1.html
  class ExternalReferences < OOXMLObject
    define_child_node(RubyXL::ExternalReference, :collection => true, :accessor => :ext_refs)
    define_element_name 'externalReferences'
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

  # http://www.schemacentral.com/sc/ooxml/e-ssml_pivotCache-1.html
  class PivotCache < OOXMLObject
    define_attribute(:cacheId, :int,    :required => true)
    define_attribute(:'r:id',  :string, :required => true)
    define_element_name 'pivotCache'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_pivotCaches-1.html
  class PivotCaches < OOXMLObject
    define_child_node(RubyXL::PivotCache, :collection => true, :accessor => :pivot_caches)
    define_element_name 'pivotCaches'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_oleSize-1.html
  class OLESize < OOXMLObject
    define_attribute(:ref, :ref, :required => true)
    define_element_name 'oleSize'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_fileRecoveryPr-1.html
  class FileRecoveryProperties < OOXMLObject
    define_attribute(:autoRecover,     :bool, :default => true)
    define_attribute(:crashSave,       :bool, :default => false)
    define_attribute(:dataExtractLoad, :bool, :default => false)
    define_attribute(:repairLoad,      :bool, :default => false)
    define_element_name 'fileRecoveryPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_webPublishing-1.html
  class WebPublishingProperties < OOXMLObject
    define_attribute(:css,              :bool,   :default => true)
    define_attribute(:thicket,          :bool,   :default => true)
    define_attribute(:longFileNames,    :bool,   :default => true)
    define_attribute(:vml,              :bool,   :default => false)
    define_attribute(:allowPng,         :bool,   :default => false)
    define_attribute(:targetScreenSize, :string, :default => '800x600',
                       :values => RubyXL::ST_TargetScreenSize)
    define_attribute(:dpi,              :int,    :default => 96)
    define_attribute(:codePage,         :int)
    define_element_name 'webPublishing'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_calcPr-1.html
  class CalculationProperties < OOXMLObject
    define_attribute(:calcId,                :int)
    define_attribute(:calcMode,              :string, :default => 'auto', :values => RubyXL::ST_CalcMode)
    define_attribute(:fullCalcOnLoad,        :bool,   :default => false)
    define_attribute(:refMode,               :string, :default => 'A1', :values => RubyXL::ST_RefMode)
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

  # http://www.schemacentral.com/sc/ooxml/e-ssml_webPublishObject-1.html
  class WebPublishObject < OOXMLObject
    define_attribute(:id,              :int,    :required => true)
    define_attribute(:divId,           :string, :required => true)
    define_attribute(:sourceObject,    :string)
    define_attribute(:destinationFile, :string, :required => true)
    define_attribute(:title,           :string)
    define_attribute(:autoRepublish,   :bool,   :default => false)
    define_element_name 'webPublishObject'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_webPublishObjects-1.html
  class WebPublishObjectContainer < OOXMLObject
    define_child_node(RubyXL::WebPublishObject, :collection => :with_count, :node_name => :web_publish_objects)
    define_element_name 'webPublishObjects'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_smartTagPr-1.html
  class SmartTagProperties < OOXMLObject
    define_attribute(:embed, :bool,   :default => false)
    define_attribute(:show,  :string, :default => 'all', :values => RubyXL::ST_SmartTagShow)
    define_element_name 'smartTagPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_smartTagType-1.html
  class SmartTagType < OOXMLObject
    define_attribute(:namespaceUri, :string)
    define_attribute(:name,         :string)
    define_attribute(:url,          :string)
    define_element_name 'smartTagType'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_smartTagTypes-1.html
  class SmartTagTypeContainer < OOXMLObject
    define_child_node(RubyXL::SmartTagType, :collection => :true, :node_name => :smart_tag_types)
    define_element_name 'smartTagTypes'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_workbook.html
  class Workbook < OOXMLTopLevelObject
    define_child_node(RubyXL::FileVersion)
    define_child_node(RubyXL::FileSharing)
    define_child_node(RubyXL::WorkbookProperties, :accessor => :workbook_properties)
    define_child_node(RubyXL::WorkbookProtection)
    define_child_node(RubyXL::WorkbookViews)
    define_child_node(RubyXL::Sheets,             :accessor => :worksheet_container)
#    ssml:functionGroups [0..1]    Function Groups
    define_child_node(RubyXL::ExternalReferences, :accessor => :ext_ref_container)
    define_child_node(RubyXL::DefinedNames,       :accessor => :defined_name_container)
    define_child_node(RubyXL::CalculationProperties)
    define_child_node(RubyXL::OLESize)
#    ssml:customWorkbookViews [0..1]    Custom Workbook Views
    define_child_node(RubyXL::PivotCaches, :accessor => :pivot_cache_container)
    define_child_node(RubyXL::SmartTagProperties)
    define_child_node(RubyXL::SmartTagTypeContainer)
    define_child_node(RubyXL::WebPublishingProperties)
    define_child_node(RubyXL::FileRecoveryProperties)
    define_child_node(RubyXL::WebPublishObjectContainer)
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

    def company
      self.document_properties.company && self.document_properties.company.value
    end

    def company=(v)
      self.document_properties.company ||= StringNode.new
      self.document_properties.company.value = v
    end

    def application
      self.document_properties.application && self.document_properties.application.value
    end

    def application=(v)
      self.document_properties.application ||= StringNode.new
      self.document_properties.application.value = v
    end

    def appversion
      self.document_properties.app_version && self.document_properties.app_version.value
    end

    def appversion=(v)
      self.document_properties.app_version ||= StringNode.new
      self.document_properties.app_version.value = v
    end

  end

end
