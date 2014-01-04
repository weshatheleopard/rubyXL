module RubyXL
  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetView-1.html
  class SheetView < OOXMLObject
    define_attribute(:tab_selected,      :tabSelected,     :int,    true)
    define_attribute(:zoom_scale,        :zoomScale,       :int,    true,  100)
    define_attribute(:zoom_scale_normal, :zoomScaleNormal, :int,    true,  100)
    define_attribute(:workbook_view_id,  :workbookViewId,  :int,    false, 0)
    define_attribute(:view,              :view,            :string, true, 
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


  class Pane
    attr_accessor :x_split, :y_split, :top_left_cell
    attr_accessor :active_pane
    VALID_PANES = %w{ bottomRight topRight bottomLeft topLeft }

    def initialize
      @x_split = @y_split = @top_left_cell = @active_pane = nil
    end

    def self.parse(node)
      pane = self.new

      pane.x_split       = RubyXL::Parser.attr_int(node, 'xSplit')
      pane.y_split       = RubyXL::Parser.attr_int(node, 'ySplit')
      pane.top_left_cell = RubyXL::Parser.attr_string(node, 'topLeftCell')
      pane.active_pane   = RubyXL::Parser.attr_string(node, 'activePane')
      pane
    end 

    def write_xml(xml)
      xml.create_element('pane', { :xSplit => @x_split, :ySplit => @y_split,
                                   :topLeftCell => @top_left_cell, :activePane => @active_pane  } )
    end

  end


  class Selection < OOXMLObject
    attr_accessor :sqref
    define_attribute(:pane,           :pane,         :string, true, nil,
                       %w{ bottomRight topRight bottomLeft topLeft })
    define_attribute(:active_cell,    :activeCell,   :string, true)
    define_attribute(:active_cell_id, :activeCellId, :int,    true) # 0-based index of @active_cell in @sqref

    def initialize
      super
      @sqref = []            # Array of references to the selected cells.
    end

    def self.parse(node)
      sel = super

      sel.active_cell    = RubyXL::Reference.new(sel.active_cell) if sel.active_cell
      sqref = RubyXL::Parser.attr_string(node, 'sqref')
      sel.sqref          = sqref.split(' ').collect{ |str| RubyXL::Reference.new(str) } if sqref
      sel
    end 

    def write_xml(xml)
      # Normally, rindex of activeCellId in sqref:
      # <selection activeCell="E12" activeCellId="9" sqref="A4 B6 C8 D10 E12 A4 B6 C8 D10 E12"/>
      if @active_cell_id.nil? && !@active_cell.nil? && @sqref.size > 1 then
        # But, things can be more complex:
        # <selection activeCell="E8" activeCellId="2" sqref="A4:B4 C6:D6 E8:F8"/>
        # Not using .reverse.each here to avoid memory reallocation.
        @sqref.each_with_index { |ref, ind| @active_cell_id = ind if ref.cover?(@active_cell) } 
      end

      attrs = prepare_attributes

      if @sqref then
        attrs[:sqref] = @sqref.join(' ')
      end

      xml.create_element('selection', attrs)
    end

  end

end
