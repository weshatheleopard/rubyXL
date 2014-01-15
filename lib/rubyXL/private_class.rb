module RubyXL
  class PrivateClass
    private

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

    #modifies fill array (copies, appends, adds color and solid attribute)
    #then styles array (copies, appends)
    def modify_fill(workbook, style_index, rgb)
      xf = workbook.cell_xfs[style_index]
      new_fill = RubyXL::Fill.new(:pattern_fill => 
                   RubyXL::PatternFill.new(:pattern_type => 'solid', :fg_color => RubyXL::Color.new(:rgb => rgb)))
      new_xf = workbook.register_new_fill(new_fill, xf)
      workbook.register_new_xf(new_xf, style_index)
    end

    #is_horizontal is true when doing horizontal alignment,
    #false when doing vertical alignment
    def modify_alignment(workbook, style_index, is_horizontal, alignment)
      old_xf = workbook.cell_xfs[style_index]

      xf = old_xf.dup
      xf.alignment ||= RubyXL::Alignment.new

      if is_horizontal then xf.alignment.horizontal = alignment
      else                  xf.alignment.vertical   = alignment
      end
      xf.apply_alignment = true

      workbook.register_new_xf(xf, style_index)
    end

    def modify_text_wrap(workbook, style_index, wrapText = false)
      old_xf = workbook.cell_xfs[style_index]

      xf = old_xf.dup
      xf.alignment ||= RubyXL::Alignment.new
      xf.alignment.wrap_text = wrapText
      xf.apply_alignment = true

      workbook.register_new_xf(xf, style_index)
    end

  end
end
