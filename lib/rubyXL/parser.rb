require 'nokogiri'
require 'tmpdir'
require 'zip'
require 'rubyXL/generic_storage'

module RubyXL

  class Parser
    def self.parse(file_path, opts = {})
      self.new(opts).parse(file_path)
    end

    # +:data_only+ allows only the sheet data to be parsed, so as to speed up parsing
    # However, using this option will result in date-formatted cells being interpreted as numbers
    def initialize(opts = {})
      @data_only = opts.is_a?(TrueClass) || opts[:data_only]
      @skip_filename_check = opts[:skip_filename_check]
    end

    def data_only
      @data_only = true
      self
    end

    def parse(xl_file_path, opts = {})
      raise 'Not .xlsx or .xlsm excel file' unless @skip_filename_check ||
                                              %w{.xlsx .xlsm}.include?(File.extname(xl_file_path))

      dir_path = File.join(File.dirname(xl_file_path), Dir::Tmpname.make_tmpname(['rubyXL', '.tmp'], nil))

      ::Zip::File.open(xl_file_path) { |zip_file|
        zip_file.each { |f|
          fpath = File.join(dir_path, f.name)
          FileUtils.mkdir_p(File.dirname(fpath))
          zip_file.extract(f, fpath) unless File.exist?(fpath)
        }
      }

      wb = RubyXL::Workbook.parse_file(dir_path)
      wb.filepath = xl_file_path

      wb.content_types = RubyXL::ContentTypes.parse_file(dir_path)
      wb.relationship_container = RubyXL::WorkbookRelationships.parse_file(dir_path)
      wb.root_relationship_container = RubyXL::RootRelationships.parse_file(dir_path)
      wb.shared_strings_container = RubyXL::SharedStringsTable.parse_file(dir_path)

      # We must always load the stylesheet because it tells us which values are actually dates/times.
      wb.stylesheet = RubyXL::Stylesheet.parse_file(dir_path)

      unless @data_only
        wb.media.load_dir(dir_path)
        wb.external_links.load_dir(dir_path)
        wb.external_links_rels.load_dir(dir_path)
        wb.drawings.load_dir(dir_path)
        wb.drawings_rels.load_dir(dir_path)
        wb.charts.load_dir(dir_path)
        wb.chart_rels.load_dir(dir_path)
        wb.printer_settings.load_dir(dir_path)
#        wb.worksheet_rels.load_dir(dir_path)
#        wb.chartsheet_rels.load_dir(dir_path)
        wb.macros.load_file(dir_path, 'vbaProject.bin')
        wb.thumbnail.load_file(dir_path, 'thumbnail.jpeg')

        wb.theme = RubyXL::Theme.parse_file(dir_path)
        wb.core_properties = RubyXL::CoreProperties.parse_file(dir_path)
        wb.document_properties = RubyXL::DocumentProperties.parse_file(dir_path)
        wb.calculation_chain = RubyXL::CalculationChain.parse_file(dir_path)
      end

      #fills out count information for each font, fill, and border
      wb.cell_xfs.each { |style|
        id = style.font_id
        wb.fonts[id].count += 1 #unless id.nil?

        id = style.fill_id
        wb.fills[id].count += 1 #unless id.nil?

        id = style.border_id
        wb.borders[id].count += 1 #unless id.nil?
      }

      wb.sheets.each_with_index { |sheet, i|
        sheet_rel = wb.relationship_container.find_by_rid(sheet.r_id)
        file_path = File.join('xl', sheet_rel.target)

        case File::basename(sheet_rel.type)
        when 'worksheet' then
          sheet_obj = RubyXL::Worksheet.parse_file(dir_path, file_path)
        when 'chartsheet' then
          sheet_obj = RubyXL::Chartsheet.parse_file(dir_path, file_path)
        end

        sheet_obj.workbook = wb
        sheet_obj.sheet_name = sheet.name
        sheet_obj.sheet_id = sheet.sheet_id
        sheet_obj.state = sheet.state

        wb.worksheets[i] = sheet_obj
        rels = SheetRelationships.parse_file(dir_path, File.join(File.dirname(file_path), '_rels', File.basename(file_path) + '.rels'))
        if rels then
          rels.sheet = sheet_obj
          sheet_obj.rels = rels
          rels.load_files(dir_path).each { |obj|
            case obj
            when RubyXL::Comments then sheet_obj.comments = obj
            end
          }

        end
      }

      FileUtils.remove_entry_secure(dir_path)

      return wb
    end

  end
end
