require 'tmpdir'
require 'date'
require 'zip'

module RubyXL
  module LegacyWorkbook
    include Enumerable

    SHEET_NAME_TEMPLATE = 'Sheet%d'
    APPLICATION = 'Microsoft Macintosh Excel'
    APPVERSION  = '12.0000'

    @@debug = nil

    def initialize(worksheets=[], filepath=nil, creator=nil, modifier=nil, created_at=nil,
                   company='', application=APPLICATION,
                   appversion=APPVERSION, date1904=0)
      super()

      # Order of sheets in the +worksheets+ array corresponds to the order of pages in Excel UI.
      # SheetId's, rId's, etc. are completely unrelated to ordering.
      @worksheets = worksheets
      add_worksheet if @worksheets.empty?

      @creator             = creator
      @modifier            = modifier
      self.date1904        = date1904 > 0

      @theme                    = RubyXL::Theme.defaults
      @shared_strings_container = RubyXL::SharedStringsTable.new
      @stylesheet               = RubyXL::Stylesheet.default
      @relationship_container   = RubyXL::WorkbookRelationships.new
      @root                     = RubyXL::WorkbookRoot.default
      @root.workbook            = self
      @root.filepath            = filepath
      @comments                 = []

      self.company         = company
      self.application     = application
      self.appversion      = appversion

      begin
        @created_at       = DateTime.parse(created_at).strftime('%Y-%m-%dT%TZ')
      rescue
        @created_at       = Time.now.strftime('%Y-%m-%dT%TZ')
      end
      @modified_at        = @created_at
    end

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

    # Return the resulting XLSX file in a stream (useful for sending over HTTP)
    def stream
      root.stream
    end

    # Save the resulting XLSX file to the specified location
    def save(filepath = nil)
      filepath ||= root.filepath

      extension = File.extname(filepath)
      unless %w{.xlsx .xlsm}.include?(extension.downcase)
        raise "Unsupported extension: #{extension} (only .xlsx and .xlsm files are supported)."
      end

      File.open(filepath, "w") { |output_file| FileUtils.copy_stream(root.stream, output_file) }

      return filepath
    end
    alias_method :write, :save

    def base_date
      if date1904 then 
        DateTime.new(1904, 1, 1)
      else
        # Subtracting one day to accomodate for erroneous 1900 leap year compatibility only for 1900 based dates
        DateTime.new(1899, 12, 31) - 1
      end
    end
    private :base_date

    def date_to_num(date)
      date && (date.ajd - base_date().ajd).to_f
    end

    def num_to_date(num)
      num && (base_date + num)
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

    private

=begin
    #fills shared strings hash, contains each unique string
    def fill_shared_strings()
      @worksheets.compact.each { |sheet|
        sheet.sheet_data.rows.each { |row|
          row.cells.each { |cell|
            if cell && cell.value && cell.datatype == RubyXL::DataType::SHARED_STRING then
              get_index(cell.value.to_s, :add_if_missing)
            end
          }
        }
      }
    end
=end

  end
end
