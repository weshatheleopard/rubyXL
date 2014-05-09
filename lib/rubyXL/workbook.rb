require 'tmpdir'
require 'zip'

module RubyXL
  module LegacyWorkbook
    include Enumerable
    attr_accessor :worksheets, :filepath, :theme,
      :external_links, :external_links_rels, :drawings, :drawings_rels,
      :charts, :chart_rels,
      :macros, :thumbnail,
      :comments, :rels_hash

    attr_accessor :stylesheet, :shared_strings_container, :calculation_chain,
                  :document_properties, :core_properties,
                  :relationship_container, :root_relationship_container, :content_types

    SHEET_NAME_TEMPLATE = 'Sheet%d'
    APPLICATION = 'Microsoft Macintosh Excel'
    APPVERSION  = '12.0000'

    def initialize(worksheets=[], filepath=nil, creator=nil, modifier=nil, created_at=nil,
                   company='', application=APPLICATION,
                   appversion=APPVERSION, date1904=0)
      super()

      # Order of sheets in the +worksheets+ array corresponds to the order of pages in Excel UI.
      # SheetId's, rId's, etc. are completely unrelated to ordering.
      @worksheets = worksheets
      add_worksheet if @worksheets.empty?

      @filepath            = filepath
      @creator             = creator
      @modifier            = modifier
      self.date1904        = date1904 > 0
      @external_links      = RubyXL::GenericStorage.new(File.join('xl', 'externalLinks'))
      @external_links_rels = RubyXL::GenericStorage.new(File.join('xl', 'externalLinks', '_rels'))
#      @drawings            = RubyXL::GenericStorage.new(File.join('xl', 'drawings'))
      @drawings_rels       = RubyXL::GenericStorage.new(File.join('xl', 'drawings', '_rels'))
      @charts              = RubyXL::GenericStorage.new(File.join('xl', 'charts'))
      @chart_rels          = RubyXL::GenericStorage.new(File.join('xl', 'charts', '_rels'))
#      @worksheet_rels      = RubyXL::GenericStorage.new(File.join('xl', 'worksheets', '_rels'))
#      @chartsheet_rels     = RubyXL::GenericStorage.new(File.join('xl', 'chartsheets', '_rels'))
      @macros              = RubyXL::GenericStorage.new('xl').binary
      @thumbnail           = RubyXL::GenericStorage.new('docProps').binary

      @theme                    = RubyXL::Theme.defaults
      @shared_strings_container = RubyXL::SharedStringsTable.new
      @stylesheet               = RubyXL::Stylesheet.default
      @document_properties      = RubyXL::DocumentProperties.new
      @core_properties          = RubyXL::CoreProperties.new
      @content_types            = RubyXL::ContentTypes.new
      @relationship_container   = RubyXL::WorkbookRelationships.new
      @root_relationship_container  = RubyXL::RootRelationships.new
      @calculation_chain            = nil
      @comments                     = []
      @rels_hash = {}

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

      new_worksheet = Worksheet.new(:workbook => self, :sheet_name => name || get_default_name)
      worksheets << new_worksheet
      new_worksheet
    end

    def each
      worksheets.each{|i| yield i}
    end


    include RubyXL::RelationshipSupport

    def related_objects
      relationship_container.workbook = root_relationship_container.workbook =
        content_types.workbook = self
      [ relationship_container, root_relationship_container, content_types ] + @worksheets
    end

    #filepath of xlsx file (including file itself)
    def write(filepath = @filepath)
      extension = File.extname(filepath)
      unless %w{.xlsx .xlsm}.include?(extension)
        raise "Only xlsx and xlsm files are supported. Unsupported extension: #{extension}"
      end

      dirpath  = File.dirname(filepath)
      temppath = File.join(dirpath, Dir::Tmpname.make_tmpname([ File.basename(filepath), '.tmp' ], nil))
      FileUtils.mkdir_p(temppath)
      zippath  = File.join(temppath, 'file.zip')

      ::Zip::File.open(zippath, ::Zip::File::CREATE) { |zipfile|

        self.rels_hash = {}
        content_types.overrides = []
        content_types.add_override(self)

        collect_related_objects.compact.each { |obj|
# puts "--> DEBUG: adding relationship to #{obj.class}"
          rels_hash[obj.class] ||= []
          rels_hash[obj.class] << obj
        }

        [ theme, stylesheet, shared_strings_container, calculation_chain, 
          document_properties, core_properties ].compact.each { |obj| 
            content_types.add_override(obj)
            obj.add_to_zip(zipfile)
          }

        [ @external_links, @external_links_rels, @macros, @thumbnail ].compact.each { |obj| obj.add_to_zip(zipfile) }

        rels_hash.each_pair { |klass, arr|
puts "--> DEBUG: saving related files of class #{klass}"
puts arr.collect{ |x| x.class }.inspect
          arr.each { |obj|
            obj.workbook = self if obj.respond_to?(:workbook=)
puts obj.class
puts "--> DEBUG:   * #{obj.xlsx_path}"
            content_types.add_override(obj)
            obj.add_to_zip(zipfile)
          }
        }

        self.add_to_zip(zipfile)
      }

      FileUtils.mv(zippath, filepath)
      FileUtils.rm_rf(temppath) if File.exist?(filepath)

      return filepath
    end

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

      unless fills[old_xf.fill_id].count == 1 && old_xf.fill_id > 2 # If the old fill is not used anymore, just replace it
        new_xf.fill_id = fills.find_index { |x| x == new_fill } # Use existing fill, if it exists
        new_xf.fill_id ||= fills.size # If this fill has never existed before, add it to collection.
      end

      fills[old_xf.fill_id].count -= 1
      new_fill.count += 1
      fills[new_xf.fill_id] = new_fill

      new_xf.apply_fill = true
      new_xf
    end

    def register_new_font(new_font, old_xf)
      new_xf = old_xf.dup

      unless fonts[old_xf.font_id].count == 1 && old_xf.font_id > 1 # If the old font is not used anymore, just replace it
        new_xf.font_id = fonts.find_index { |x| x == new_font } # Use existing font, if it exists
        new_xf.font_id ||= fonts.size # If this font has never existed before, add it to collection.
      end

      fonts[old_xf.font_id].count -= 1
      new_font.count += 1
      fonts[new_xf.font_id] = new_font

      new_xf.apply_font = true
      new_xf
    end

    def register_new_xf(new_xf, old_style_index)
      new_xf_id = cell_xfs.find_index { |xf| xf == new_xf } # Use existing XF, if it exists
      new_xf_id ||= cell_xfs.size # If this XF has never existed before, add it to collection.

      cell_xfs[old_style_index].count -= 1
      new_xf.count += 1
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
      xf.alignment = RubyXL::Alignment.new(:apply_alignment => true,
                                           :horizontal => is_horizontal ? alignment : nil, 
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

      unless borders[old_xf.border_id].count == 1 && old_xf.border_id > 0 # If the old border not used anymore, just replace it
        new_xf.border_id = borders.find_index { |x| x == new_border } # Use existing border, if it exists
        new_xf.border_id ||= borders.size # If this border has never existed before, add it to collection.
      end

      borders[old_xf.border_id].count -= 1
      new_border.count += 1
      borders[new_xf.border_id] = new_border

      new_xf.apply_border = true

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
