module RubyXL
  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetView-1.html
  class SheetView < OOXMLObject
    define_attribute(:tab_selected,      :tabSelected,     :int)
    define_attribute(:zoom_scale,        :zoomScale,       :int,    false,  100)
    define_attribute(:zoom_scale_normal, :zoomScaleNormal, :int,    false,  100)
    define_attribute(:workbook_view_id,  :workbookViewId,  :int,    :required, 0)
    define_attribute(:view,              :view,            :string, false, 
                       %w{ normal pageBreakPreview pageLayout })

    attr_accessor :pane, :selections

    def initialize
      @pane = nil
      @selections = []
      super
    end

    def self.parse(node)
      sheetview = super

      node.element_children.each { |child_node|
        case child_node.name
        when 'pane' then sheetview.pane = RubyXL::Pane.parse(child_node)
        when 'selection' then sheetview.selections << RubyXL::Selection.parse(child_node)
        else raise "Node type #{child_node.name} not implemented"
        end
      }

      sheetview
    end 

    def write_xml(xml)
      node = xml.create_element('sheetView', prepare_attributes)
      node << pane.write_xml(xml) if @pane
      @selections.each { |sel| node << sel.write_xml(xml) }
      node
    end

  end


  class Pane < OOXMLObject
    define_attribute(:x_split,       :xSplit,      :int)
    define_attribute(:y_split,       :ySplit,      :int)
    define_attribute(:top_left_cell, :topLeftCell, :string)
    define_attribute(:active_pane,   :activePane,  :string, false, nil,
                       %w{ bottomRight topRight bottomLeft topLeft })
    define_element_name 'pane'
  end


  class Selection < OOXMLObject
    define_attribute(:pane,           :pane,         :string, false, nil,
                       %w{ bottomRight topRight bottomLeft topLeft })
    define_attribute(:active_cell,    :activeCell,   :string)
    define_attribute(:active_cell_id, :activeCellId, :int)              # 0-based index of @active_cell in @sqref
    define_attribute(:sqref,          :sqref,        :sqref, :required) # Array of references to the selected cells.
    define_element_name 'selection'

    def self.parse(node)
      sel = super

      sel.active_cell = RubyXL::Reference.new(sel.active_cell) if sel.active_cell
      sel
    end 

    def before_write_xml
      # Normally, rindex of activeCellId in sqref:
      # <selection activeCell="E12" activeCellId="9" sqref="A4 B6 C8 D10 E12 A4 B6 C8 D10 E12"/>
      if @active_cell_id.nil? && !@active_cell.nil? && @sqref.size > 1 then
        # But, things can be more complex:
        # <selection activeCell="E8" activeCellId="2" sqref="A4:B4 C6:D6 E8:F8"/>
        # Not using .reverse.each here to avoid memory reallocation.
        @sqref.each_with_index { |ref, ind| @active_cell_id = ind if ref.cover?(@active_cell) } 
      end
    end

  end

end
