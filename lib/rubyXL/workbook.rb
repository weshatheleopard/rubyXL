require File.expand_path(File.join(File.dirname(__FILE__),'writer','content_types_writer'))
require File.expand_path(File.join(File.dirname(__FILE__),'writer','root_rels_writer'))
require File.expand_path(File.join(File.dirname(__FILE__),'writer','app_writer'))
require File.expand_path(File.join(File.dirname(__FILE__),'writer','core_writer'))
require File.expand_path(File.join(File.dirname(__FILE__),'writer','theme_writer'))
require File.expand_path(File.join(File.dirname(__FILE__),'writer','workbook_rels_writer'))
require File.expand_path(File.join(File.dirname(__FILE__),'writer','workbook_writer'))
require File.expand_path(File.join(File.dirname(__FILE__),'writer','styles_writer'))
require File.expand_path(File.join(File.dirname(__FILE__),'writer','shared_strings_writer'))
require File.expand_path(File.join(File.dirname(__FILE__),'writer','worksheet_writer'))
require 'rubyXL/zip'
require 'date'

module RubyXL
  class Workbook
    include Enumerable
    attr_accessor :worksheets, :filepath, :creator, :modifier, :created_at,
      :modified_at, :company, :application, :appversion, :num_fmts, :num_fmts_hash, :fonts, :fills,
      :borders, :cell_xfs, :cell_style_xfs, :cell_styles, :shared_strings, :calc_chain,
      :num_strings, :size, :date1904, :external_links, :style_corrector, :drawings,
      :worksheet_rels, :printer_settings, :macros, :colors, :shared_strings_XML, :defined_names, :column_lookup_hash


    APPLICATION = 'Microsoft Macintosh Excel'
    APPVERSION  = '12.0000'
    SHEET_NAME = 'Sheet1'
    def initialize(worksheets=[], filepath=nil, creator=nil, modifier=nil, created_at=nil,
                   company='', application=APPLICATION,
                   appversion=APPVERSION, date1904=0)
      if worksheets.nil? || worksheets.empty?
        @worksheets       = [Worksheet.new(self,SHEET_NAME)]
      else
        @worksheets       = worksheets
      end
      @filepath           = filepath
      @creator            = creator
      @modifier           = modifier
      @company            = company
      @application        = application
      @appversion         = appversion
      @num_fmts           = nil
      @num_fmts_hash      = nil
      @fonts              = nil
      @fills              = nil
      @borders            = nil
      @cell_xfs           = nil
      @cell_style_xfs     = nil
      @cell_styles        = nil
      @shared_strings     = nil
      @calc_chain         = nil #unnecessary?
      @num_strings        = 0 #num strings total
      @size               = 0 #num strings in shared_strings array
      @date1904           = date1904 > 0
      @external_links     = nil
      @style_corrector    = nil
      @drawings           = nil
      @worksheet_rels     = nil
      @printer_settings   = nil
      @macros             = nil
      @colors             = nil
      @shared_strings_XML = nil
      @defined_names      = nil
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

    # allows easier access to worksheets
    def [](worksheet)
      return worksheets[worksheet]
    end

    def each
      worksheets.each{|i| yield i}
    end

    def num_fmts_by_id
      
      return @num_fmts_hash unless @num_fmts_hash.nil?
      if num_fmts
        @num_fmts_hash={}
        num_fmts[:numFmt].each do |num_fmt|
          @num_fmts_hash[num_fmt[:attributes][:numFmtId]]=num_fmt
        end
        @num_fmts_hash
      else
        {}
      end
    end

    #filepath of xlsx file (including file itself)
    def write(filepath=@filepath)
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
      FileUtils.mkdir_p(File.join(dirpath,zippath))
      Zip::ZipFile.open(zippath, Zip::ZipFile::CREATE) do |zipfile|
        writer = Writer::ContentTypesWriter.new(dirpath,self)
        zipfile.get_output_stream('[Content_Types].xml') {|f| f.puts(writer.write())}

        writer = Writer::RootRelsWriter.new(dirpath,self)
        zipfile.get_output_stream(File.join('_rels','.rels')) {|f| f.puts(writer.write())}

        writer = Writer::AppWriter.new(dirpath,self)
        zipfile.get_output_stream(File.join('docProps','app.xml')) {|f| f.puts(writer.write())}

        writer = Writer::CoreWriter.new(dirpath,self)
        zipfile.get_output_stream(File.join('docProps','core.xml')) {|f| f.puts(writer.write())}

        writer = Writer::ThemeWriter.new(dirpath,self)
        zipfile.get_output_stream(File.join('xl','theme','theme1.xml')) {|f| f.puts(writer.write())}

        writer = Writer::WorkbookRelsWriter.new(dirpath,self)
        zipfile.get_output_stream(File.join('xl','_rels','workbook.xml.rels')) {|f| f.puts(writer.write())}

        writer = Writer::WorkbookWriter.new(dirpath,self)
        zipfile.get_output_stream(File.join('xl','workbook.xml')) {|f| f.puts(writer.write())}

        writer = Writer::StylesWriter.new(dirpath,self)
        zipfile.get_output_stream(File.join('xl','styles.xml')) {|f| f.puts(writer.write())}

        unless @shared_strings.nil?
          writer = Writer::SharedStringsWriter.new(dirpath,self)
          zipfile.get_output_stream(File.join('xl','sharedStrings.xml')) {|f| f.puts(writer.write())}
        end

        #preserves external links (exactly, no modification allowed)
        unless @external_links.nil?
          #-1 because of rels
          1.upto(@external_links.size-1) do |i|
            zipfile.get_output_stream(
              File.join('xl','externalLinks',"externalLink#{i}.xml")) {|f|
                f.puts(@external_links[i])
              }
          end
          @external_links['rels'].each_index do |i|
            unless @external_links['rels'][i].nil?
              zipfile.get_output_stream(
                File.join('xl','externalLinks','_rels',"externalLink#{i}.xml.rels")) {|f|
                  f.puts(@external_links['rels'][i])
                }
            end
          end
        end

        #preserves drawings (exactly, no modification allowed)
        unless @drawings.nil?
          1.upto(@drawings.size) do |i|
            zipfile.get_output_stream(
            File.join('xl','drawings',"vmlDrawing#{i}.vml")) {|f|
              f.puts(@drawings[i])
            }
          end
        end

        unless @printer_settings.nil?
          1.upto(@printer_settings.size) do |i|
            zipfile.get_output_stream(
            File.join('xl','printerSettings',"printerSettings#{i}.bin")) {|f|
              f.puts(@printer_settings[i])
            }
          end
        end

        unless @worksheet_rels.nil?
          1.upto(@worksheet_rels.size) do |i|
            zipfile.get_output_stream(
            File.join('xl','worksheets','_rels',"sheet#{i}.xml.rels")) {|f|
              f.puts(@worksheet_rels[i])
            }
          end
        end

        unless @macros.nil?
          zipfile.get_output_stream(File.join('xl','vbaProject.bin')) {|f| f.puts(@macros)}
        end

        @worksheets.each_with_index do |sheet,i|
          writer = Writer::WorksheetWriter.new(dirpath,self,i)
          zipfile.get_output_stream(File.join('xl','worksheets',"sheet#{i+1}.xml")) {|f| f.puts(writer.write())}
        end
      end

      FileUtils.cp(zippath,File.join(dirpath,filename+".#{extension}"))
      FileUtils.cp(File.join(dirpath,filename+".#{extension}"),filepath)
      if File.exist?(filepath)
        FileUtils.rm_rf(dirpath)
      end
      return filepath
    end

    def date_to_num(date)
      return nil if date.nil?
      if @date1904
        compare_date = DateTime.parse('December 31, 1903')
      else
        compare_date = DateTime.parse('December 31, 1899')
      end
      # add one day to compare date for erroneous 1900 leap year compatibility
      date.ajd + 1 - compare_date.ajd
    end

    def num_to_date(num)
      return nil if num.nil?
      if @date1904
        compare_date = DateTime.parse('December 31, 1903')
      else
        compare_date = DateTime.parse('December 31, 1899')
      end
      # subtract one day to compare date for erroneous 1900 leap year compatibility
      compare_date - 1 + num
    end

    def date_num_fmt?(num_fmt)
      @num_fmt_date_hash ||= {}
      if @num_fmt_date_hash[num_fmt].nil?
        @num_fmt_date_hash[num_fmt] = is_date_format?(num_fmt)
      end
      return @num_fmt_date_hash[num_fmt]
    end

     def is_date_format?(num_fmt)
      skip_chars = ['$', '-', '+', '/', '(', ')', ':', ' ']
      num_chars = ['0', '#', '?']
      non_date_formats = ['0.00E+00', '##0.0E+0', 'General', 'GENERAL', 'general', '@']
      date_chars = ['y','m','d','h','s']

      state = 0
      s = ''
      num_fmt.split(//).each do |c|
        if state == 0
          if c == '"'
            state = 1
          elsif ['\\', '_', '*'].include?(c)
            state = 2
          elsif skip_chars.include?(c)
            next
          else
            s << c
          end
        elsif state == 1
          if c == '"'
            state = 0
          end
        elsif state == 2
          state = 0
        end
      end
      s.gsub!(/\[[^\]]*\]/, '')
      if non_date_formats.include?(s)
        return false
      end
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
      if @fills[xf_attributes[:fillId]].nil? || @fills[xf_attributes[:fillId]][:fill].nil? || @fills[xf_attributes[:fillId]][:fill][:patternFill].nil? || @fills[xf_attributes[:fillId]][:fill][:patternFill][:fgColor].nil?
        'ffffff' #white
      else
        @fills[xf_attributes[:fillId]][:fill][:patternFill][:fgColor][:attributes][:rgb]
      end
    end


    private

    # Do not change. Excel requires that some of these styles be default,
    # and will simply assume that the 0 and 1 indexed fonts are the default values.
    def fill_styles()
      @fonts = {
                 '0' => {
                          :font => {
                                     :sz   => { :attributes => { :val => 10 } },
                                     :name => { :attributes => { :val => "Verdana" } }
                                   },
                          :count=>1
                        },
                 '1' => {
                          :font => {
                                     :sz   => { :attributes => { :val => 8 } },
                                     :name => { :attributes => { :val => "Verdana" } }
                                   },
                          :count=>0
                        }
               }

      @fills = {
                 '0' => {
                          :fill => {
                                     :patternFill => { :attributes => { :patternType => "none" } }
                                   },
                  :count=>1} ,
                 '1' => {
                          :fill => {
                                     :patternFill => { :attributes => { :patternType => "gray125" } }
                                   },
                          :count=>0
                        }
               }

      @borders = {
                   '0' => {
                            :border => {
                                         :left => { },
                                         :right => { },
                                         :top => { },
                                         :bottom => { },
                                         :diagonal => { }
                                       },
                            :count => 1 #count = how many styles reference it
                          }
                 }

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
      if @shared_strings.nil?
        string_hash = {}
        string_index = 0
        @num_strings = 0
        #fill hash for shared strings
        @worksheets.each do |sheet|
          unless sheet.nil?
            sheet.sheet_data.each do |row|
              row.each do |cell|
                unless cell.nil? || cell.value.nil?
                  #if string not already seen, add it to hash
                  if cell.datatype == 's'
                    if string_hash[cell.value.to_s].nil?
                      string_hash[string_index]=cell.value.to_s
                      string_hash[cell.value.to_s]=string_index
                      string_index += 1
                    end
                    @num_strings += 1
                  end
                end
              end
            end
          end
        end

        if string_hash.empty?
          @shared_strings = nil
        else
          @shared_strings = string_hash
          @size = string_index
        end
      end
    end

    def validate_before_write
      ## TODO CHECK IF STYLE IS OK if not raise
    end
  end
end
