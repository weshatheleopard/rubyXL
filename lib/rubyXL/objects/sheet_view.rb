module RubyXL

=begin
<sheetView tabSelected="1" workbookViewId="0">
<pane xSplit="3270" ySplit="1200" topLeftCell="D4" activePane="bottomRight"/>
<selection activeCell="B3" sqref="B3"/>
<selection pane="topRight" activeCell="E3" sqref="E3"/>
<selection pane="bottomLeft" activeCell="B5" sqref="B5"/>
<selection pane="bottomRight" activeCell="E2" activeCellId="3" sqref="E5 B5 B2 E2"/>
</sheetView>
=end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetView-1.html
  class SheetView
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

      node.element_children.each { |c|
        case c.name
        when 'pane' then sheetview.pane = RubyXL::Pane.parse(c)
        when 'selection' then # TODO#
        else raise "Node type #{c.name} not implemented"
        end
      }

      sheetview
    end 

    def write_xml(xml)
      attrs = { :workbookViewId  => @workbook_view_id } 
      attrs[:tabSelected]     = @tab_selected      unless @tab_selected.nil?
      attrs[:view]            = @view              unless @view.nil?
      attrs[:zoomScale]       = @zoom_scale        unless @zoom_scale.nil?
      attrs[:zoomScaleNormal] = @zoom_scale_normal unless @zoom_scale_normal.nil?

      node = xml.create_element('sheetView', attrs)
      node << pane.write_xml(xml) if @pane
      
      @selections.each { |sel|
        node << sel.write_xml(xml)
      }
      
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

end
