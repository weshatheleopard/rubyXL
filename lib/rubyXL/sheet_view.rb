module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetView-1.html
  class SheetView
    attr_accessor :tab_selected, :zoom_scale, :zoom_scale_normal, :workbook_view_id, :view

    def initialize
      @tab_selected     = @view = nil
      @zoom_scale       = @zoom_scale_normal = 100
      @workbook_view_id = 0
    end

    def self.parse(node)
      sheetview = self.new

      sheetview.tab_selected      = RubyXL::Parser.attr_int(node, 'tabSelected')
      sheetview.zoom_scale        = RubyXL::Parser.attr_int(node, 'zoomScale')
      sheetview.zoom_scale_normal = RubyXL::Parser.attr_int(node, 'zoomScaleNormal')
      sheetview.workbook_view_id  = RubyXL::Parser.attr_int(node, 'workbookViewId')
      # Valid values: 'normal', 'pageBreakPreview', 'pageLayout'
      sheetview.view              = RubyXL::Parser.attr_string(node, 'view')
      sheetview
    end 


    def write_xml(xml)
      attrs = { :workbookViewId  => @workbook_view_id } 
      attrs[:tabSelected]     = @tabSelected unless @tabSelected.nil?
      attrs[:view]            = @view        unless @view.nil?
      attrs[:zoomScale]       = @zoom_scale  unless @zoom_scale.nil?
      attrs[:zoomScaleNormal] = @zoom_scale_normal unless @zoom_scale_normal.nil?

      xml << xml.create_element('sheetView', attrs)
    end

  end

end
