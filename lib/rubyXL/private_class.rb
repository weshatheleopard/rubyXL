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

    def validate_text_wrap(wrap)
      raise 'Only true or false are valid wraps' unless wrap.is_a?(FalseClass) || wrap.is_a?(TrueClass)
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
      workbook.fonts.each_with_index {|f, font_id|
        if (f.italic == font.italic &&
          f.bold == font.bold &&
          f.underlined == font.underlined &&
          f.strikethrough == font.strikethrough &&
          f.name == font.name &&
          f.size == font.size &&
          f.color == font.color) # TODO: need to write proper comparison
          return font_id
        end
      }
      return nil
    end

    # Helper method to modify the font color
    def modify_font_color(font, font_color)
      font.color = RubyXL::Color.new if font.color.nil?
      font.color.rgb = font_color.to_s
      return font
    end

    # Helper method to modify the font's italics settings
    def modify_font_italics(font, italicized)
      font.italic = italicized
      return font
    end

    # Helper method to modify the font's bold settings
    def modify_font_bold(font, bolded)
      font.bold = bolded
      return font
    end

    # Helper method to modify the font's underline settings
    def modify_font_underline(font, underlined)
      font.underlined = underlined
      return font
    end

    # Helper method to modify the font's strikethrough settings
    def modify_font_strikethrough(font, struckthrough)
      font.strikethrough = struckthrough
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
        workbook.fonts[font_id].count += 1
        workbook.fonts[old_font_id].count -= 1
      elsif workbook.fonts[old_font_id].count > 1 || old_font_id == 0
        font_id = workbook.fonts.size
        workbook.fonts[font_id] = font
        workbook.fonts[font_id].count = 1
        workbook.fonts[old_font_id].count -= 1
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
      xf = workbook.get_style_attributes(workbook.get_style(style_index))

      new_fill_id = fill_id = xf[:fillId]

      fill = workbook.fills[fill_id]

      # If the current fill is used in more than one cell, we need to create a copy;
      # otherwise we can modify it in place (with the exception of Special Fills #1 and #0)
      if fill.count > 1 || fill_id == 0 || fill_id == 1
        new_fill_id = workbook.fills.size
        fill.count -= 1

        workbook.cell_xfs[:xf] = [workbook.cell_xfs[:xf]] unless workbook.cell_xfs[:xf].is_a?(Array)
        workbook.cell_xfs[:xf] << deep_copy({ :attributes => xf })

        xf = workbook.get_style_attributes(workbook.cell_xfs[:xf].last)
        xf[:fillId] = new_fill_id
        xf[:applyFill] = 1
        workbook.cell_xfs[:attributes][:count] += 1
        style_index = workbook.cell_xfs[:xf].size - 1
      end
 
      new_fill = RubyXL::Fill.new(:pattern_fill => 
                   RubyXL::PatternFill.new(:pattern_type => 'solid', :fg_color => RubyXL::Color.new(:rgb => rgb)))
      new_fill.count = 1
      workbook.fills[new_fill_id] = new_fill
        
      return style_index
    end

    def modify_border(workbook, style_index)
      xf = workbook.get_style_attributes(workbook.get_style(style_index))

      new_border_id = border_id = xf[:borderId]

      border = workbook.borders[border_id]

      # If the current border is used in more than one cell, we need to create a copy;
      # otherwise we can modify it in place (with the exception of Special Border #0)
      if border.count > 1 || border_id == 0
        new_border_id = workbook.borders.size
        border.count -= 1

        workbook.cell_xfs[:xf] = [workbook.cell_xfs[:xf]] unless workbook.cell_xfs[:xf].is_a?(Array)
        workbook.cell_xfs[:xf] << deep_copy({ :attributes => xf })

        xf = workbook.get_style_attributes(workbook.cell_xfs[:xf].last)
        xf[:borderId] = border_id
        xf[:applyBorder] = 1
        workbook.cell_xfs[:attributes][:count] += 1
        style_index =  workbook.cell_xfs[:xf].size-1
      end

      new_border = RubyXL::Border.new()
      new_border.count = 1
      workbook.borders[new_border_id] = new_border
        
      return style_index
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

    def modify_text_wrap(workbook, style_index, wrapText=0)
      old_xf_obj = workbook.get_style(style_index)

      xf_obj = deep_copy(old_xf_obj)

      if xf_obj[:alignment].nil? || xf_obj[:alignment][:attributes].nil?
        xf_obj[:alignment] = {:attributes=>{:horizontal=>nil, :vertical=>nil, :wrapText => nil}}
      end

      xf_obj[:alignment][:attributes][:wrapText] = (wrapText ? "1" : "0")
      workbook.cell_xfs[:xf] = [workbook.cell_xfs[:xf]] unless workbook.cell_xfs[:xf].is_a?(Array)
      workbook.cell_xfs[:xf] << deep_copy(xf_obj)

      xf = workbook.get_style_attributes(workbook.cell_xfs[:xf].last)
      xf[:applyAlignment] = '1'
      workbook.cell_xfs[:attributes][:count] += 1
      workbook.cell_xfs[:xf].size-1
    end

    #returns non-shallow copy of hash
    def deep_copy(hash)
      Marshal.load(Marshal.dump(hash))
    end

  end
end
