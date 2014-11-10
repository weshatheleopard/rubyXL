module RubyXL
  module WorkbookConvenienceMethods
    SHEET_NAME_TEMPLATE = 'Sheet%d'

    # Finds worksheet by its name or numerical index
    def [](ind)
      case ind
      when Integer then worksheets[ind]
      when String  then worksheets.find { |ws| ws.sheet_name == ind }
      end
    end

    # Create new simple worksheet and add it to the workbook worksheets
    #
    # @param [String] The name for the new worksheet
    def add_worksheet(name = nil)
      if name.nil? then
        n = 0

        begin
          name = SHEET_NAME_TEMPLATE % (n += 1)
        end until self[name].nil?
      end

      new_worksheet = Worksheet.new(:workbook => self, :sheet_name => name)
      worksheets << new_worksheet
      new_worksheet
    end

    def each
      worksheets.each{ |i| yield i }
    end

    def date1904
      workbook_properties && workbook_properties.date1904
    end

    def date1904=(v)
      self.workbook_properties ||= RubyXL::WorkbookProperties.new
      workbook_properties.date1904 = v
    end

    def company
      root.document_properties.company && root.document_properties.company.value
    end

    def company=(v)
      root.document_properties.company ||= StringNode.new
      root.document_properties.company.value = v
    end

    def application
      root.document_properties.application && root.document_properties.application.value
    end

    def application=(v)
      root.document_properties.application ||= StringNode.new
      root.document_properties.application.value = v
    end

    def appversion
      root.document_properties.app_version && root.document_properties.app_version.value
    end

    def appversion=(v)
      root.document_properties.app_version ||= StringNode.new
      root.document_properties.app_version.value = v
    end

    def creator
      root.core_properties.creator
    end

    def creator=(v)
      root.core_properties.creator = v
    end

    def modifier
      root.core_properties.modifier
    end

    def modifier=(v)
      root.core_properties.modifier = v
    end

    def created_at
      root.core_properties.created_at
    end

    def created_at=(v)
      root.core_properties.created_at = v
    end

    def modified_at
      root.core_properties.modified_at
    end

    def modified_at=(v)
      root.core_properties.modified_at = v
    end

    def cell_xfs # Stylesheet should be pre-filled with defaults on initialize()
      stylesheet.cell_xfs
    end

    def fonts # Stylesheet should be pre-filled with defaults on initialize()
      stylesheet.fonts
    end

    def fills # Stylesheet should be pre-filled with defaults on initialize()
      stylesheet.fills
    end

    def borders # Stylesheet should be pre-filled with defaults on initialize()
      stylesheet.borders
    end

    def get_fill_color(xf)
      fill = fills[xf.fill_id]
      pattern = fill && fill.pattern_fill
      color = pattern && pattern.fg_color
      color && color.rgb || 'ffffff'
    end

    def register_new_fill(new_fill, old_xf)
      new_xf = old_xf.dup
      new_xf.apply_fill = true
      new_xf.fill_id = fills.find_index { |x| x == new_fill } # Use existing fill, if it exists
      new_xf.fill_id ||= fills.size # If this fill has never existed before, add it to collection.
      fills[new_xf.fill_id] = new_fill
      new_xf
    end

    def register_new_font(new_font, old_xf)
      new_xf = old_xf.dup
      new_xf.apply_font = true
      new_xf.font_id = fonts.find_index { |x| x == new_font } # Use existing font, if it exists
      new_xf.font_id ||= fonts.size # If this font has never existed before, add it to collection.
      fonts[new_xf.font_id] = new_font
      new_xf
    end

    def register_new_xf(new_xf, old_style_index)
      new_xf_id = cell_xfs.find_index { |xf| xf == new_xf } # Use existing XF, if it exists
      new_xf_id ||= cell_xfs.size # If this XF has never existed before, add it to collection.
      cell_xfs[new_xf_id] = new_xf
      new_xf_id
    end

    def modify_text_wrap(style_index, wrap = false)
      xf = cell_xfs[style_index].dup
      xf.alignment = RubyXL::Alignment.new(:wrap_text => wrap, :apply_alignment => true)
      register_new_xf(xf, style_index)
    end

    def modify_alignment(style_index, is_horizontal, alignment)
      xf = cell_xfs[style_index].dup
      xf.apply_alignment = true
      xf.alignment = RubyXL::Alignment.new(:horizontal => is_horizontal ? alignment : nil,
                                           :vertical   => is_horizontal ? nil : alignment)
      register_new_xf(xf, style_index)
    end

    def modify_fill(style_index, rgb)
      xf = cell_xfs[style_index].dup
      new_fill = RubyXL::Fill.new(:pattern_fill =>
                   RubyXL::PatternFill.new(:pattern_type => 'solid',
                                           :fg_color => RubyXL::Color.new(:rgb => rgb)))
      new_xf = register_new_fill(new_fill, xf)
      register_new_xf(new_xf, style_index)
    end

    def modify_border(style_index, direction, weight)
      old_xf = cell_xfs[style_index].dup
      new_border = borders[old_xf.border_id].dup
      new_border.set_edge_style(direction, weight)

      new_xf = old_xf.dup
      new_xf.apply_border = true

      new_xf.border_id = borders.find_index { |x| x == new_border } # Use existing border, if it exists
      new_xf.border_id ||= borders.size # If this border has never existed before, add it to collection.
      borders[new_xf.border_id] = new_border

      register_new_xf(new_xf, style_index)
    end

  end


  module WorksheetConvenienceMethods

    def get_row_fill(row = 0)
      (row = sheet_data.rows[row]) && row.get_fill_color
    end

    def get_row_font_name(row = 0)
      (font = row_font(row)) && font.get_name
    end

    def get_row_font_size(row = 0)
      (font = row_font(row)) && font.get_size
    end

    def get_row_font_color(row = 0)
      font = row_font(row)
      color = font && font.color
      color && (color.rgb || '000000')
    end

    def is_row_italicized(row = 0)
      (font = row_font(row)) && font.is_italic
    end

    def is_row_bolded(row = 0)
      (font = row_font(row)) && font.is_bold
    end

    def is_row_underlined(row = 0)
      (font = row_font(row)) && font.is_underlined
    end

    def is_row_struckthrough(row = 0)
      (font = row_font(row)) && font.is_strikethrough
    end

    def get_row_height(row = 0)
      validate_workbook
      validate_nonnegative(row)
      return nil unless row_exists(row)
      row = sheet_data.rows[row]
      row && row.ht || 13
    end

    def get_row_horizontal_alignment(row = 0)
      return get_row_alignment(row, true)
    end

    def get_row_vertical_alignment(row = 0)
      return get_row_alignment(row, false)
    end

    def get_row_border_top(row = 0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_row_border` instead."
      return get_row_border(row, :top)
    end

    def get_row_border_left(row = 0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_row_border` instead."
      return get_row_border(row, :left)
    end

    def get_row_border_right(row = 0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_row_border` instead."
      return get_row_border(row, :right)
    end

    def get_row_border_bottom(row = 0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_row_border` instead."
      return get_row_border(row, :bottom)
    end

    def get_row_border_diagonal(row = 0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_row_border` instead."
      return get_row_border(row, :diagonal)
    end

    def get_column_font_name(col = 0)
      font = column_font(col)
      font && font.get_name
    end

    def get_column_font_size(col = 0)
      font = column_font(col)
      font && font.get_size
    end

    def get_column_font_color(col = 0)
      font = column_font(col)
      font && (font.get_rgb_color || '000000')
    end

    def is_column_italicized(col = 0)
      font = column_font(col)
      font && font.is_italic
    end

    def is_column_bolded(col = 0)
      font = column_font(col)
      font && font.is_bold
    end

    def is_column_underlined(col = 0)
      font = column_font(col)
      font && font.is_underlined
    end

    def is_column_struckthrough(col = 0)
      font = column_font(col)
      font && font.is_strikethrough
    end

    # Get raw column width value as stored in the file
    def get_column_width_raw(column_index = 0)
      validate_workbook
      validate_nonnegative(column_index)

      range = cols.locate_range(column_index)
      range && range.width
    end

    # Get column width measured in number of digits, as per
    # http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column%28v=office.14%29.aspx
    def get_column_width(column_index = 0)
      width = get_column_width_raw(column_index)
      return RubyXL::ColumnRange::DEFAULT_WIDTH if width.nil?
      (width - (5.0 / RubyXL::Font::MAX_DIGIT_WIDTH)).round
    end

    def get_column_fill(col=0)
      validate_workbook
      validate_nonnegative(col)
      return nil unless column_exists(col)

      @workbook.get_fill_color(get_col_xf(col))
    end

    def get_column_horizontal_alignment(col=0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_column_alignment` instead."
      get_column_alignment(col, :horizontal)
    end

    def get_column_vertical_alignment(col=0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_column_alignment` instead."
      get_column_alignment(col, :vertical)
    end

    def get_column_border_top(col=0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_column_border` instead."
      get_column_border(col, :top)
    end

    def get_column_border_left(col=0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_column_border` instead."
      get_column_border(col, :left)
    end

    def get_column_border_right(col=0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_column_border` instead."
      get_column_border(col, :right)
    end

    def get_column_border_bottom(col=0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_column_border` instead."
      get_column_border(col, :bottom)
    end

    def get_column_border_diagonal(col=0)
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_column_border` instead."
      get_column_border(col, :diagonal)
    end

    def change_column_horizontal_alignment(column_index, alignment = 'center')
      change_column_alignment(column_index, alignment, true)
    end

    def change_column_vertical_alignment(column_index, alignment = 'center')
      change_column_alignment(column_index, alignment, false)
    end

    def change_column_border_top(column_index, weight = 'thin')
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `change_column_border` instead."
      change_column_border(column_index, :top, weight)
    end

    def change_column_border_left(column_index, weight = 'thin')
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `change_column_border` instead."
      change_column_border(column_index, :left, weight)
    end

    def change_column_border_right(column_index, weight = 'thin')
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `change_column_border` instead."
      change_column_border(column_index, :right, weight)
    end

    def change_column_border_bottom(column_index, weight = 'thin')
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `change_column_border` instead."
      change_column_border(column_index, :bottom, weight)
    end

    def change_column_border_diagonal(column_index, weight = 'thin')
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `change_column_border` instead."
      change_column_border(column_index, :diagonal, weight)
    end

    def change_column_border(column_index, direction, weight)
      validate_workbook
      ensure_cell_exists(0, column_index)

      cols.get_range(column_index).style_index = @workbook.modify_border(get_col_style(column_index), direction, weight)

      sheet_data.rows.each { |row|
        c = row.cells[column_index]
        c.change_border(direction, weight) unless c.nil?
      }
    end

  end


  module CellConvenienceMethods

    def change_horizontal_alignment(alignment = 'center')
      validate_worksheet
      self.style_index = workbook.modify_alignment(self.style_index, true, alignment)
    end

    def change_vertical_alignment(alignment = 'center')
      validate_worksheet
      self.style_index = workbook.modify_alignment(self.style_index, false, alignment)
    end

    def change_text_wrap(wrap = false)
      validate_worksheet
      self.style_index = workbook.modify_text_wrap(self.style_index, wrap)
    end

    def change_border(direction, weight)
      validate_worksheet
      self.style_index = workbook.modify_border(self.style_index, direction, weight)
    end

    def change_border_top(weight = 'thin')
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `change_border` instead."
      change_border(:top, weight)
    end

    def change_border_left(weight = 'thin')
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `change_border` instead."
      change_border(:left, weight)
    end

    def change_border_right(weight = 'thin')
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `change_border` instead."
      change_border(:right, weight)
    end

    def change_border_bottom(weight = 'thin')
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `change_border` instead."
      change_border(:bottom, weight)
    end

    def change_border_diagonal(weight = 'thin')
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `change_border` instead."
      change_border(:diagonal, weight)
    end

    def border_top()
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_border` instead."
      return get_border(:top)
    end

    def border_left()
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_border` instead."
      return get_border(:left)
    end

    def border_right()
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_border` instead."
      return get_border(:right)
    end

    def border_bottom()
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_border` instead."
      return get_border(:bottom)
    end

    def border_diagonal()
      warn "[DEPRECATION] `#{__method__}` is deprecated.  Please use `get_border` instead."
      return get_border(:diagonal)
    end

  end

end
