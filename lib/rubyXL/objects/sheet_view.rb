module RubyXL
  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetView-1.html
  class SheetView
    include RubyXL::XMLhelper

    attr_accessor :tab_selected, :zoom_scale, :zoom_scale_normal, :workbook_view_id, :view,
                  :pane, :selections

    def initialize
      @tab_selected     = @view = nil
      @zoom_scale       = @zoom_scale_normal = 100
      @workbook_view_id = 0
      @pane = nil
      @selections = []
    end

    def self.parse(node)
      sheetview = self.new

      sheetview.tab_selected      = RubyXL::Parser.attr_int(node, 'tabSelected')
      sheetview.zoom_scale        = RubyXL::Parser.attr_int(node, 'zoomScale')
      sheetview.zoom_scale_normal = RubyXL::Parser.attr_int(node, 'zoomScaleNormal')
      sheetview.workbook_view_id  = RubyXL::Parser.attr_int(node, 'workbookViewId')
      # Valid values: 'normal', 'pageBreakPreview', 'pageLayout'
      sheetview.view              = RubyXL::Parser.attr_string(node, 'view')

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
      @attrs = { :workbookViewId  => @workbook_view_id } 
      attr_optional(:tabSelected,     @tab_selected)
      attr_optional(:view,            @view)
      attr_optional(:zoomScale,       @zoom_scale)
      attr_optional(:zoomScaleNormal, @zoom_scale_normal)

      node = xml.create_element('sheetView', @attrs)
      node << pane.write_xml(xml) if @pane
      @selections.each { |sel| node << sel.write_xml(xml) }
      node
    end

  end


  class Pane
    attr_accessor :x_split, :y_split, :top_left_cell, :active_pane

    def initialize
      @x_split = @y_split = @top_left_cell = @active_pane = nil
    end

    def self.parse(node)
      pane = self.new

      pane.x_split       = RubyXL::Parser.attr_int(node, 'xSplit')
      pane.y_split       = RubyXL::Parser.attr_int(node, 'ySplit')
      pane.top_left_cell = RubyXL::Parser.attr_string(node, 'topLeftCell')
      # Valid values: [ 'bottomRight', 'topRight', 'bottomLeft', 'topLeft' ]
      pane.active_pane   = RubyXL::Parser.attr_string(node, 'activePane')
      pane
    end 

    def write_xml(xml)
      xml.create_element('pane', { :xSplit => @x_split, :ySplit => @y_split,
                                   :topLeftCell => @top_left_cell, :activePane => @active_pane  } )
    end

  end


  class Selection
    include XMLhelper
    attr_accessor :pane, :active_cell, :active_cell_id, :sqref

    def initialize
      @pane = nil
      @sqref = []            # Array of references to the selected cells.
      @active_cell = nil
      @active_cell_id = nil  # index of @active_cell in @sqref
    end

    def self.parse(node)
      sel = self.new

      sel.pane           = RubyXL::Parser.attr_string(node, 'pane')
      sel.active_cell    = RubyXL::Parser.attr_string(node, 'activeCell')
      sel.active_cell_id = RubyXL::Parser.attr_int(node, 'activeCellId')
      sqref              = RubyXL::Parser.attr_string(node, 'sqref')
      sel.sqref          = sqref && sqref.split(' ')
      sel
    end 

    def write_xml(xml)
      @attrs = {}
      attr_optional(:pane,         @pane)
      attr_optional(:activeCell,   @active_cell)

      # Normally, rindex of activeCellId in sqref:
      # <selection activeCell="E12" activeCellId="9" sqref="A4 B6 C8 D10 E12 A4 B6 C8 D10 E12"/>
      # but, can be more complex:
      # <selection activeCell="E8" activeCellId="2" sqref="A4:B4 C6:D6 E8:F8"/>
      # TODO: update activeCellId detection once Reference class is implemented.
      attr_optional(:activeCellId, @active_cell_id || @sqref.rindex(@active_cell)) 

      if @sqref then
        @attrs[:sqref] = @sqref.join(' ')
      end

      xml.create_element('selection', @attrs)
    end

  end

end
