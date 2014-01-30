require 'rubyXL/objects/ooxml_object'

module RubyXL
  # http://www.schemacentral.com/sc/ooxml/e-ssml_pane-1.html
  class Pane < OOXMLObject
    define_attribute(:xSplit,      :int)
    define_attribute(:ySplit,      :int)
    define_attribute(:topLeftCell, :string)
    define_attribute(:activePane,  :string, :default => 'topLeft',
                       :values => %w{ bottomRight topRight bottomLeft topLeft })
    define_attribute(:state,       :string, :default=> 'split',
                       :values => %w{ split frozen frozenSplit })
    define_element_name 'pane'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_selection-1.html
  class Selection < OOXMLObject
    define_attribute(:pane,         :string,
                       :values => %w{ bottomRight topRight bottomLeft topLeft })
    define_attribute(:activeCell,   :ref)
    define_attribute(:activeCellId, :int)   # 0-based index of @active_cell in @sqref
    define_attribute(:sqref,        :sqref) # Array of references to the selected cells.
    define_element_name 'selection'

    def before_write_xml
      # Normally, rindex of activeCellId in sqref:
      # <selection activeCell="E12" activeCellId="9" sqref="A4 B6 C8 D10 E12 A4 B6 C8 D10 E12"/>
      if @active_cell_id.nil? && !@active_cell.nil? && @sqref.size > 1 then
        # But, things can be more complex:
        # <selection activeCell="E8" activeCellId="2" sqref="A4:B4 C6:D6 E8:F8"/>
        # Not using .reverse.each here to avoid memory reallocation.
        @sqref.each_with_index { |ref, ind| @active_cell_id = ind if ref.cover?(@active_cell) } 
      end
      true
    end
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetView-1.html
  class SheetView < OOXMLObject
    define_attribute(:windowProtection,         :bool, :default => false)
    define_attribute(:showFormulas,             :bool, :default => false)
    define_attribute(:showGridLines,            :bool, :default => true)
    define_attribute(:showRowColHeaders,        :bool, :default => true)
    define_attribute(:showZeros,                :bool, :default => true)
    define_attribute(:rightToLeft,              :bool, :default => false)
    define_attribute(:tabSelected,              :bool, :default => false)
    define_attribute(:showRuler,                :bool, :default => true)
    define_attribute(:showOutlineSymbols,       :bool, :default => true)
    define_attribute(:defaultGridColor,         :bool, :default => true)
    define_attribute(:showWhiteSpace,           :bool, :default => true)
    define_attribute(:view,                     :string, :default => 'normal',
                       :values => %w{ normal pageBreakPreview pageLayout })
    define_attribute(:topLeftCell,              :ref)
    define_attribute(:colorId,                  :int, :default => 64)
    define_attribute(:zoomScale,                :int, :default => 100)
    define_attribute(:zoomScaleNormal,          :int, :default => 0)
    define_attribute(:zoomScaleSheetLayoutView, :bool, :default => 0)
    define_attribute(:zoomScalePageLayoutView,  :bool, :default => 0)
    define_attribute(:workbookViewId,           :int, :required => true, :default => 0 )
    define_child_node(RubyXL::Pane)
    define_child_node(RubyXL::Selection, :collection => true, :accessor => :selections )
    define_element_name 'sheetView'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetViews-3.html
  class SheetViews < OOXMLObject
    define_child_node(RubyXL::SheetView, :collection => true, :accessor => :sheet_views)
    define_element_name 'sheetViews'
  end

end
