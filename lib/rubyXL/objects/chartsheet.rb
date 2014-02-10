require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/simple_types'
require 'rubyXL/objects/extensions'
require 'rubyXL/objects/relationships'
require 'rubyXL/objects/sheet_common'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetProtection-4.html
  class ChartsheetProtection < OOXMLObject
    define_attribute(:password, :string)
    define_attribute(:content,  :bool, :default => false)
    define_attribute(:objects,  :bool, :default => false)
    define_element_name 'sheetProtection'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetPr-4.html
  class ChartsheetProperties < OOXMLObject
    define_attribute(:published,                         :bool, :default => true)
    define_attribute(:codeName,                          :string)
    define_child_node(RubyXL::Color, :node_name => :tabColor)
    define_element_name 'sheetPr'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_pageSetup-5.html
  class ChartsheetPageSetup < OOXMLObject
    define_attribute(:paperSize,          :int,    :default => 1)
    define_attribute(:firstPageNumber,    :int,    :default => 1)
    define_attribute(:orientation,        :string, :default => 'default',
                       :values => RubyXL::ST_Orientation)
    define_attribute(:usePrinterDefaults, :bool,   :default => true)
    define_attribute(:blackAndWhite,      :bool,   :default => false)
    define_attribute(:draft,              :bool,   :default => false)
    define_attribute(:useFirstPageNumber, :bool,   :default => false)
    define_attribute(:horizontalDpi,      :int,    :default => 600)
    define_attribute(:verticalDpi,        :int,    :default => 600)
    define_attribute(:copies,             :int,    :default => 1)
    define_attribute(:'r:id',             :string)
    define_element_name 'pageSetup'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetView-2.html
  class ChartsheetView < OOXMLObject
    define_attribute(:tabSelected,    :bool,  :default => false)
    define_attribute(:zoomScale,      :int,   :default => 100)
    define_attribute(:workbookViewId, :int,   :required => true, :default => 0 )
    define_attribute(:zoomToFit,      :bool,  :default => false)
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'sheetView'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetViews-4.html
  class ChartsheetViewContainer < OOXMLObject
    define_child_node(RubyXL::ChartsheetView, :collection => true, :accessor => :sheet_views)
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'sheetViews'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_chartsheet.html
  class Chartsheet < OOXMLTopLevelObject
    define_child_node(RubyXL::ChartsheetProperties)
    define_child_node(RubyXL::ChartsheetViewContainer, :accessor => :sheet_view_container)
    define_child_node(RubyXL::ChartsheetProtection)
    define_child_node(RubyXL::CustomSheetViews, :accessor => :custom_sheet_view_container)
    define_child_node(RubyXL::PageMargins)
    define_child_node(RubyXL::ChartsheetPageSetup)
    define_child_node(RubyXL::HeaderFooterSettings)
    define_child_node(RubyXL::RID, :node_name => :drawing)
    define_child_node(RubyXL::RID, :node_name => :legacyDrawing)
    define_child_node(RubyXL::RID, :node_name => :legacyDrawingHF)
    define_child_node(RubyXL::RID, :node_name => :picture)
#    ssml:webPublishItems [0..1]    Web Publishing Items
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'chartsheet'
    set_namespaces('xmlns'       => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                   'xmlns:r'     => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
  end

end
