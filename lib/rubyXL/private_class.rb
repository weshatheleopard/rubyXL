module RubyXL
  class PrivateClass
    private

    #validate and modify methods
    def validate_horizontal_alignment(alignment)
      if alignment.to_s == '' || alignment == 'center' || alignment == 'distributed' || alignment == 'justify' || alignment == 'left' || alignment == 'right'
        return true
      end
      raise 'Only center, distributed, justify, left, and right are valid horizontal alignments'
    end

    def validate_vertical_alignment(alignment)
      if alignment.to_s == '' || alignment == 'center' || alignment == 'distributed' || alignment == 'justify' || alignment == 'top' || alignment == 'bottom'
        return true
      end
      raise 'Only center, distributed, justify, top, and bottom are valid vertical alignments'
    end

    def validate_border(weight)
      if weight.to_s == '' || weight == 'thin' || weight == 'thick' || weight == 'hairline' || weight == 'medium'
        return true
      end
      raise 'Border weights must only be "hairline", "thin", "medium", or "thick"'
    end

    def validate_nonnegative(row_or_col)
      if row_or_col < 0
        raise 'Row and Column arguments must be nonnegative'
      end
    end

    # This method checks to see if there is an equivalent font that exists
    def find_font(workbook, font)
      workbook.fonts.each {|font_id, f|
        if f[:font][:i] == font[:i] &&
          f[:font][:b] == font[:b] &&
          f[:font][:u] == font[:u] &&
          f[:font][:strike] == font[:strike] &&
          f[:font][:name][:attributes][:val] == font[:name][:attributes][:val] &&
          f[:font][:sz][:attributes][:val] == font[:sz][:attributes][:val] &&
          (f[:font][:color] && f[:font][:color][:attributes][:rgb]) == (font[:color] && font[:color][:attributes][:rgb])
          return font_id
        end
      }
      return nil
    end

    # Helper method to modify the font color
    def modify_font_color(font, font_color)
      if font[:color].nil?
        font[:color] = {:attributes => {:rgb => ''}}
      end
      font[:color][:attributes][:rgb] = font_color.to_s
      return font
    end

    # Helper method to modify the font's italics settings
    def modify_font_italics(font, italicized)
      if italicized
        font[:i] = {}
      else
        font[:i] = nil
      end
      return font
    end

    # Helper method to modify the font's bold settings
    def modify_font_bold(font, bolded)
      if bolded
        font[:b] = {}
      else
        font[:b] = nil
      end
      return font
    end

    # Helper method to modify the font's underline settings
    def modify_font_underline(font, underlined)
      if underlined
        font[:u] = {}
      else
        font[:u] = nil
      end
      return font
    end

    # Helper method to modify the font's strikethrough settings
    def modify_font_strikethrough(font, struckthrough)
      if struckthrough
        font[:strike] = {}
      else
        font[:strike] = nil
      end
      return font
    end

    # Determines if font exists
    # If yes, return id of existing font
    # If no, appends font to font array
    def modify_font(workbook, font, old_font_id)
      font_id = old_font_id
      existing_font_id = find_font(workbook, font)
      if !existing_font_id.nil?
        font_id = existing_font_id
        workbook.fonts[font_id][:count] += 1
        workbook.fonts[old_font_id][:count] -= 1
      elsif workbook.fonts[old_font_id.to_s][:count] > 1 || old_font_id == '0'
        font_id = workbook.fonts.size.to_s
        workbook.fonts[font_id] = {}
        workbook.fonts[font_id][:font] = font
        workbook.fonts[font_id][:count] = 1
        workbook.fonts[old_font_id][:count] -= 1
      else
        workbook.fonts[font_id][:font] = font
      end
      return font_id
    end

    # This method checks to see if there is an equivalent xf that exists
    def find_xf(workbook, xf)
      workbook.cell_xfs[:xf].each_with_index {|xfs, index|
        if xfs[:attributes][:borderId] == xf[:borderId] &&
          xfs[:attributes][:xfId] == xf[:xfId] &&
          xfs[:attributes][:fillId] == xf[:fillId] &&
          xfs[:attributes][:numFmtId] == xf[:numFmtId] &&
          xfs[:attributes][:fontId] == xf[:fontId]
          return index
        end
      }
      return nil
    end

    # Determines if xf exists
    # If yes, return id of existing xf
    # If no, appends xf to xf array
    def modify_xf(workbook, xf)
      existing_xf_id = find_xf(workbook, xf)
      if !existing_xf_id.nil?
        xf_id = existing_xf_id
      else
        if workbook.cell_xfs[:xf].is_a?Array
          workbook.cell_xfs[:xf] << {:attributes=>xf}
        else
          workbook.cell_xfs[:xf] = [workbook.cell_xfs[:xf], {:attributes=>xf}]
        end
        xf[:applyFont] = '1'
        workbook.cell_xfs[:attributes][:count] += 1
        xf_id = workbook.cell_xfs[:xf].size - 1
      end
      return xf_id
    end

    #modifies fill array (copies, appends, adds color and solid attribute)
    #then styles array (copies, appends)
    def modify_fill(workbook, style_index, rgb)
      xf_obj = workbook.get_style(style_index)
      xf = workbook.get_style_attributes(xf_obj)
      #modify fill array
      fill_id = xf[:fillId]

      fill = workbook.fills[fill_id.to_s][:fill]
      if workbook.fills[fill_id.to_s][:count] > 1 || fill_id == 0 || fill_id == 1
        old_size = workbook.fills.size.to_s
        workbook.fills[old_size] = {}
        workbook.fills[old_size][:fill] = deep_copy(fill)
        workbook.fills[old_size][:count] = 1
        workbook.fills[fill_id.to_s][:count] -= 1

        change_wb_fill(workbook, old_size,rgb)

        #modify styles array
        fill_id = old_size
        if workbook.cell_xfs[:xf].is_a?Array
          workbook.cell_xfs[:xf] << deep_copy({:attributes=>xf})
        else
          workbook.cell_xfs[:xf] = [workbook.cell_xfs[:xf], deep_copy({:attributes=>xf})]
        end
        xf = workbook.get_style_attributes(workbook.cell_xfs[:xf].last)
        xf[:fillId] = fill_id
        xf[:applyFill] = '1'
        workbook.cell_xfs[:attributes][:count] += 1
        return workbook.cell_xfs[:xf].size-1
      else
        change_wb_fill(workbook, fill_id.to_s,rgb)
        return style_index
      end
    end

    def modify_border(workbook, style_index)
      xf_obj = workbook.get_style(style_index)
      xf = workbook.get_style_attributes(xf_obj)

      border_id = Integer(xf[:borderId])
      border = workbook.borders[border_id.to_s][:border]
      if workbook.borders[border_id.to_s][:count] > 1 || border_id == 0 || border_id == 1
        old_size = workbook.borders.size.to_s
        workbook.borders[old_size] = {}
        workbook.borders[old_size][:border] = deep_copy(border)
        workbook.borders[old_size][:count] = 1

        border_id = old_size

        if workbook.cell_xfs[:xf].is_a?Array
          workbook.cell_xfs[:xf] << deep_copy(xf_obj)
        else
          workbook.cell_xfs[:xf] = [workbook.cell_xfs[:xf], deep_copy(xf_obj)]
        end

        xf = workbook.get_style_attributes(workbook.cell_xfs[:xf].last)
        xf[:borderId] = border_id
        xf[:applyBorder] = '1'
        workbook.cell_xfs[:attributes][:count] += 1
        return workbook.cell_xfs[:xf].size-1
      else
        return style_index
      end
    end

    #is_horizontal is true when doing horizontal alignment,
    #false when doing vertical alignment
    def modify_alignment(workbook, style_index, is_horizontal, alignment)
      old_xf_obj = workbook.get_style(style_index)

      xf_obj = deep_copy(old_xf_obj)

      if xf_obj[:alignment].nil? || xf_obj[:alignment][:attributes].nil?
        xf_obj[:alignment] = {:attributes=>{:horizontal=>nil, :vertical=>nil}}
      end

      if is_horizontal
        xf_obj[:alignment][:attributes][:horizontal] = alignment.to_s
      else
        xf_obj[:alignment][:attributes][:vertical] = alignment.to_s
      end

      if workbook.cell_xfs[:xf].is_a?Array
        workbook.cell_xfs[:xf] << deep_copy(xf_obj)
      else
        workbook.cell_xfs[:xf] = [workbook.cell_xfs[:xf], deep_copy(xf_obj)]
      end

      xf = workbook.get_style_attributes(workbook.cell_xfs[:xf].last)
      xf[:applyAlignment] = '1'
      workbook.cell_xfs[:attributes][:count] += 1
      workbook.cell_xfs[:xf].size-1
    end

    #returns non-shallow copy of hash
    def deep_copy(hash)
      Marshal.load(Marshal.dump(hash))
    end

    def change_wb_fill(workbook, fill_index, rgb)
      if workbook.fills[fill_index][:fill][:patternFill][:fgColor].nil?
        workbook.fills[fill_index][:fill][:patternFill][:fgColor] = {:attributes => {:rgb => ''}}
      end
      workbook.fills[fill_index][:fill][:patternFill][:fgColor][:attributes][:rgb] = rgb

      #previously none, doesn't show fill
      workbook.fills[fill_index][:fill][:patternFill][:attributes][:patternType] = 'solid'
    end

  end
end
