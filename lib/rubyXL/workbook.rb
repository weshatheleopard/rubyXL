require 'rubyXL/writer/generic_writer'
require 'rubyXL/writer/content_types_writer'
require 'rubyXL/writer/root_rels_writer'
require 'rubyXL/writer/app_writer'
require 'rubyXL/writer/core_writer'
require 'rubyXL/writer/theme_writer'
require 'rubyXL/writer/workbook_rels_writer'
require 'rubyXL/writer/workbook_writer'
require 'rubyXL/writer/styles_writer'
require 'rubyXL/writer/shared_strings_writer'
require 'rubyXL/writer/worksheet_writer'
require 'rubyXL/zip'
require 'rubyXL/shared_strings'
require 'date'

module RubyXL
  class Workbook
    include Enumerable
    attr_accessor :worksheets, :filepath, :creator, :modifier, :created_at,
      :modified_at, :company, :application, :appversion, :num_fmts, :num_fmts_hash, :fonts, :fills,
      :borders, :cell_xfs, :cell_style_xfs, :cell_styles, :calc_chain, :theme,
      :date1904, :external_links, :external_links_rels, :style_corrector,
      :drawings, :drawings_rels, :charts, :chart_rels,
      :worksheet_rels, :printer_settings, :macros, :colors, :shared_strings_XML, :defined_names, :column_lookup_hash

    attr_reader :shared_strings

    APPLICATION = 'Microsoft Macintosh Excel'
    APPVERSION  = '12.0000'

    def initialize(worksheets=[], filepath=nil, creator=nil, modifier=nil, created_at=nil,
                   company='', application=APPLICATION,
                   appversion=APPVERSION, date1904=0)

      # Order of sheets in the +worksheets+ array corresponds to the order of pages in Excel UI.
      # SheetId's, rId's, etc. are completely unrelated to ordering.
      @worksheets = worksheets || []

      @worksheets << Worksheet.new(self) if @worksheets.empty?

      @filepath           = filepath
      @creator            = creator
      @modifier           = modifier
      @company            = company
      @application        = application
      @appversion         = appversion
      @num_fmts           = nil
      @num_fmts_hash      = nil
      @fonts              = []
      @fills              = nil
      @borders            = []
      @cell_xfs           = nil
      @cell_style_xfs     = nil
      @cell_styles        = nil
      @shared_strings     = RubyXL::SharedStrings.new
      @calc_chain         = nil #unnecessary?
      @date1904           = date1904 > 0
      @external_links     = RubyXL::GenericStorage.new(File.join('xl', 'externalLinks'))
      @external_links_rels= RubyXL::GenericStorage.new(File.join('xl', 'externalLinks', '_rels'))
      @style_corrector    = nil
      @drawings           = RubyXL::GenericStorage.new(File.join('xl', 'drawings'))
      @drawings_rels      = RubyXL::GenericStorage.new(File.join('xl', 'drawings', '_rels'))
      @charts             = RubyXL::GenericStorage.new(File.join('xl', 'charts'))
      @chart_rels         = RubyXL::GenericStorage.new(File.join('xl', 'charts', '_rels'))
      @worksheet_rels     = RubyXL::GenericStorage.new(File.join('xl', 'worksheets', '_rels'))
      @theme              = RubyXL::GenericStorage.new(File.join('xl', 'theme'))
      @printer_settings   = RubyXL::GenericStorage.new(File.join('xl', 'printerSettings')).binary
      @macros             = RubyXL::GenericStorage.new('xl').binary
      @colors             = {}
      @shared_strings_XML = nil
      @defined_names      = []
      @column_lookup_hash = {}

      begin
        @created_at       = DateTime.parse(created_at).strftime('%Y-%m-%dT%TZ')
      rescue
        t = Time.now
        @created_at       = t.strftime('%Y-%m-%dT%TZ')
      end
      @modified_at        = @created_at

      fill_styles()
      fill_shared_strings()
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
      new_worksheet = Worksheet.new(self, name)
      worksheets << new_worksheet
      new_worksheet
    end

    def each
      worksheets.each{|i| yield i}
    end

    def num_fmts_by_id
      return @num_fmts_hash unless @num_fmts_hash.nil?

      @num_fmts_hash = {}

      if num_fmts then
        num_fmts[:numFmt].each { |num_fmt|
          @num_fmts_hash[num_fmt[:attributes][:numFmtId]] = num_fmt
        }
      end

      @num_fmts_hash
    end

    #filepath of xlsx file (including file itself)
    def write(filepath = @filepath)
      validate_before_write

      if !(filepath =~ /(.+)\.xls(x|m)/)
        raise "Only xlsx and xlsm files are supported. Unsupported type for file: #{filepath}"
      end

      dirpath = ''
      extension = 'xls'

      if(filepath =~ /((.|\s)*)\.xls(x|m)$/)
        dirpath = $1.to_s()
        extension += $3.to_s
      end

      filename = ''

      if(filepath =~ /\/((.|\s)*)\/((.|\s)*)\.xls(x|m)$/)
        filename = $3.to_s()
      end

      #creates zip file, writes each type of file to zip folder
      #zips package and renames it to xlsx.
      zippath = File.join(dirpath, filename + '.zip')
      File.unlink(zippath) if File.exists?(zippath)
      FileUtils.mkdir_p(dirpath)

      Zip::File.open(zippath, Zip::File::CREATE) do |zipfile|
        [ Writer::ContentTypesWriter, Writer::RootRelsWriter, Writer::AppWriter, 
          Writer::CoreWriter, Writer::ThemeWriter, Writer::WorkbookRelsWriter,
          Writer::WorkbookWriter, Writer::StylesWriter ].each { |writer_class|
          writer_class.new(self).add_to_zip(zipfile)
        }
       
        unless @shared_strings.empty?
          Writer::SharedStringsWriter.new(self).add_to_zip(zipfile)
        end

        @external_links.add_to_zip(zipfile)
        @external_links_rels.add_to_zip(zipfile)
        @drawings.add_to_zip(zipfile)
        @drawings_rels.add_to_zip(zipfile)
        @charts.add_to_zip(zipfile)
        @chart_rels.add_to_zip(zipfile)
        @printer_settings.add_to_zip(zipfile)
        @worksheet_rels.add_to_zip(zipfile)
        @macros.add_to_zip(zipfile)
        @worksheets.each_index { |i| Writer::WorksheetWriter.new(self, i).add_to_zip(zipfile) }
      end

      full_file_path = File.join(dirpath, "#{filename}.#{extension}")
      FileUtils.cp(zippath, full_file_path)
      FileUtils.cp(full_file_path, filepath)

      if File.exist?(filepath)
        FileUtils.rm_rf(dirpath)
      end
      return filepath
    end

    def base_date
      if @date1904 then
        Date.new(1904, 1, 1)
      else
        # Subtracting one day to accomodate for erroneous 1900 leap year compatibility only for 1900 based dates
        Date.new(1899, 12, 31) - 1
      end
    end
    private :base_date

    def date_to_num(date)
      date && (date.ajd - base_date().ajd).to_i
    end

    def num_to_date(num)
      num && (base_date + num)
    end

    def date_num_fmt?(num_fmt)
      @num_fmt_date_hash ||= {}
      if @num_fmt_date_hash[num_fmt].nil?
        @num_fmt_date_hash[num_fmt] = is_date_format?(num_fmt)
      end
      return @num_fmt_date_hash[num_fmt]
    end

    def is_date_format?(num_fmt)
      num_fmt = num_fmt.downcase
      skip_chars = ['$', '-', '+', '/', '(', ')', ':', ' ']
      num_chars = ['0', '#', '?']
      non_date_formats = ['0.00e+00', '##0.0e+0', 'general', '@']
      date_chars = ['y','m','d','h','s']

      state = 0
      s = ''

      num_fmt.split(//).each do |c|
        case state 
        when 0 then
          if c == '"'
            state = 1
          elsif ['\\', '_', '*'].include?(c)
            state = 2
          elsif skip_chars.include?(c)
            next
          else
            s << c
          end
        when 1 then
          state = 0 if c == '"'
        when 2 then
          state = 0
        end
      end

      s.gsub!(/\[[^\]]*\]/, '')

      return false if non_date_formats.include?(s)

      separator = ';'
      got_sep = 0
      date_count = 0
      num_count = 0

      s.split(//).each do |c|
        if date_chars.include?(c)
          date_count += 1
        elsif num_chars.include?(c)
          num_count += 1
        elsif c == separator
          got_sep = 1
        end
      end

      if date_count > 0 && num_count == 0
        return true
      elsif num_count > 0 && date_count == 0
        return false
      elsif date_count
        # ambiguous result
      elsif got_sep == 0
        # constant result
      end

      return date_count > num_count
    end

    #gets style object from style array given index
    def get_style(style_index)
      if !@cell_xfs[:xf].is_a?Array
        @cell_xfs[:xf] = [@cell_xfs[:xf]]
      end

      xf_obj = @cell_xfs[:xf]
      if xf_obj.is_a?Array
        xf_obj = xf_obj[Integer(style_index)]
      end
      xf_obj
    end

    #gets attributes of above style object
    #necessary because can take the form of hash or array,
    #based on odd behavior of Nokogiri
    def get_style_attributes(xf_obj)
      if xf_obj.is_a?Array
        xf = xf_obj[1]
      else
        xf = xf_obj[:attributes]
      end
    end

    def get_fill_color(xf_attributes)
      fill = @fills[xf_attributes[:fillId]]
      return 'ffffff' if fill.nil? || fill.fg_color.nil?
      fill.fg_color.rgb
    end


    private

    # Do not change. Excel requires that some of these styles be default,
    # and will simply assume that the 0 and 1 indexed fonts are the default values.
    def fill_styles()

      f0 = RubyXL::Font.new
      f1 = RubyXL::Font.new
      f0.name = f1.name = 'Verdana'
      f0.size = 10
      f1.size = 8
      f0.count = 1

      @fonts = [ f0, f1 ]

      @fills = [ RubyXL::PatternFill.new(:pattern_type => 'none'),
                 RubyXL::PatternFill.new(:pattern_type => 'gray125') ]

      @borders = [ RubyXL::Border.new ]

      @cell_style_xfs = {
                        :attributes => {
                                         :count => 1
                                       },
                        :xf => {
                                 :attributes => { :numFmtId => 0, :fontId => 0, :fillId => 0, :borderId => 0 }
                               }
                      }
      @cell_xfs = {
                        :attributes => {
                                         :count => 1
                                       },
                        :xf => {
                                 :attributes => { :numFmtId => 0, :fontId => 0, :fillId => 0, :borderId => 0, :xfId => 0 }
                               }
                      }
      @cell_styles = {
                      :cellStyle => {
                                      :attributes => { :builtinId=>0, :name=>"Normal", :xfId=>0 }
                                    },
                      :attributes => { :count => 1 }
                    }
    end


    #fills shared strings hash, contains each unique string
    def fill_shared_strings()
      @worksheets.compact.each { |sheet|
        sheet.sheet_data.each { |row|
          row.each { |cell|
            if cell && cell.value && cell.datatype == RubyXL::Cell::SHARED_STRING then
              get_index(cell.value.to_s, :add_if_missing)
            end
          }
        }
      }
    end

    def validate_before_write
      ## TODO CHECK IF STYLE IS OK if not raise
    end
  end
end
