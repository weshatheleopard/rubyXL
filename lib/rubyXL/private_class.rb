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

    # This method checks to see if there is an equivalent xf that exists
    def find_xf(workbook, xf)
      workbook.cell_xfs.each_with_index { |xfs, index| return index if xfs == xf }
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
        xf.apply_font = true
        xf_id = workbook.cell_xfs.size - 1
      end
      return xf_id
    end

    #modifies fill array (copies, appends, adds color and solid attribute)
    #then styles array (copies, appends)
    def modify_fill(workbook, style_index, rgb)
      xf = workbook.cell_xfs[style_index]

      new_fill_id = fill_id = xf.fill_id

      fill = workbook.fills[fill_id]

      # If the current fill is used in more than one cell, we need to create a copy;
      # otherwise we can modify it in place (with the exception of Special Fills #1 and #0)
      if fill.count > 1 || fill_id == 0 || fill_id == 1
        fill.count -= 1
        new_fill_id = workbook.fills.size
        style_index = workbook.cell_xfs.size
        workbook.cell_xfs << RubyXL::XF.new(:fill_id => new_fill_id, :apply_fill => true)
      end
 
      new_fill = RubyXL::Fill.new(:pattern_fill => 
                   RubyXL::PatternFill.new(:pattern_type => 'solid', :fg_color => RubyXL::Color.new(:rgb => rgb)))
      new_fill.count = 1
      workbook.fills[new_fill_id] = new_fill
        
      style_index
    end

    #is_horizontal is true when doing horizontal alignment,
    #false when doing vertical alignment
    def modify_alignment(workbook, style_index, is_horizontal, alignment)
      old_xf_obj = workbook.cell_xfs[style_index]

      xf= old_xf_obj.dup
      xf.alignment ||= RubyXL::Alignment.new

      if is_horizontal
        xf.alignment.horizontal = alignment
      else
        xf.alignment.vertical = alignment
      end
      xf.apply_alignment = true

      new_index = workbook.cell_xfs.size
      workbook.cell_xfs[new_index] = xf
      new_index
    end

    def modify_text_wrap(workbook, style_index, wrapText = false)
      old_xf_obj = workbook.cell_xfs[style_index]

      xf_obj = old_xf_obj.dup
      xf_obj.alignment ||= RubyXL::Alignment.new
      xf_obj.alignment.wrap_text = wrapText
      xf.apply_alignment true

      new_index = workbook.cell_xfs.size
      workbook.cell_xfs[new_index] = xf
      new_index
    end

  end
end
