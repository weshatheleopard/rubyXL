module RubyXL
  module WorkbookConvenienceMethods
    def each(&block)
      worksheets.each(&block)
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
      color &&= color.get_rgb(self)
      (color && color.to_s) || 'ffffff'
    end

    def register_new_fill(new_fill, old_xf)
      new_xf = old_xf.dup
      new_xf.apply_fill = true
      new_xf.fill_id = fills.find_index { |x| x == new_fill } # Reuse existing fill, if it exists
      new_xf.fill_id ||= fills.size # If this fill has never existed before, add it to collection.
      fills[new_xf.fill_id] = new_fill
      new_xf
    end

    def register_new_font(new_font, old_xf)
      new_xf = old_xf.dup
      new_xf.font_id = fonts.find_index { |x| x == new_font } # Reuse existing font, if it exists
      new_xf.font_id ||= fonts.size # If this font has never existed before, add it to collection.
      fonts[new_xf.font_id] = new_font
      new_xf.apply_font = true
      new_xf
    end

    def register_new_xf(new_xf)
      new_xf_id = cell_xfs.find_index { |xf| xf == new_xf } # Reuse existing XF, if it exists
      new_xf_id ||= cell_xfs.size # If this XF has never existed before, add it to collection.
      cell_xfs[new_xf_id] = new_xf
      new_xf_id
    end

    def modify_alignment(style_index)
      old_xf = cell_xfs[style_index || 0]
      new_xf = old_xf.dup
      new_xf.alignment = if old_xf.alignment then
                           old_xf.alignment.dup
                         else
                           RubyXL::Alignment.new
                         end

      yield(new_xf.alignment)
      new_xf.apply_alignment = true

      register_new_xf(new_xf)
    end

    def modify_fill(style_index, rgb)
      xf = cell_xfs[style_index || 0].dup
      new_fill = RubyXL::Fill.new(:pattern_fill =>
                                                   RubyXL::PatternFill.new(:pattern_type => 'solid',
                                                                           :fg_color     => RubyXL::Color.new(:rgb => rgb)))
      register_new_xf(register_new_fill(new_fill, xf))
    end

    def modify_border(style_index, direction, weight)
      xf = cell_xfs[style_index || 0].dup
      new_border = borders[xf.border_id || 0].dup

      edge = new_border.send(direction)
      new_border.send("#{direction}=", edge.dup) if edge

      new_border.set_edge_style(direction, weight)

      xf.border_id = borders.find_index { |x| x == new_border } # Reuse existing border, if it exists
      xf.border_id ||= borders.size # If this border has never existed before, add it to collection.
      borders[xf.border_id] = new_border
      xf.apply_border = true

      register_new_xf(xf)
    end

    def modify_border_color(style_index, direction, color)
      xf = cell_xfs[style_index || 0].dup
      new_border = borders[xf.border_id || 0].dup
      new_border.set_edge_color(direction, color)

      xf.border_id = borders.find_index { |x| x == new_border } # Reuse existing border, if it exists
      xf.border_id ||= borders.size # If this border has never existed before, add it to collection.
      borders[xf.border_id] = new_border
      xf.apply_border = true

      register_new_xf(xf)
    end

    # Calculate password hash from string for use in 'password' fields.
    # https://www.openoffice.org/sc/excelfileformat.pdf
    def password_hash(pwd)
      hsh = 0
      pwd.reverse.each_char { |c|
        hsh ^= c.ord
        hsh <<= 1
        hsh -= 0x7fff if hsh > 0x7fff
      }

      (hsh ^ pwd.length ^ 0xCE4B).to_s(16)
    end

    def define_new_name(name, reference)
      self.defined_names ||= RubyXL::DefinedNames.new
      self.defined_names << RubyXL::DefinedName.new({ :name => name, :reference => reference })
    end

    def get_defined_name(name)
      self.defined_names && self.defined_names.find { |n| n.name == name }
    end

    def title
      root.core_properties.dc_title && root.core_properties.dc_title.value
    end

    def title=(v)
      root.core_properties.dc_title = v && RubyXL::StringNode.new(:value => v)
    end
  end

  RubyXL::Workbook.include RubyXL::WorkbookConvenienceMethods # ruby 2.1 compat
end
