require 'nokogiri'
require 'tmpdir'
require 'zip'
require 'rubyXL/generic_storage'

module RubyXL

  class Parser
    def self.parse(file_path, opts = {})
      self.new(opts).parse(file_path)
    end

    # +keep_tempfiles_on_error+ prevents temp files to be erased if exception was thrown during parsing.
    # +:data_only+ allows only the sheet data to be parsed, so as to speed up parsing.
    # However, using this option will result in date-formatted cells being interpreted as numbers.
    def initialize(opts = {})
      @data_only = opts.is_a?(TrueClass) || opts[:data_only]
      @skip_filename_check = opts[:skip_filename_check]
      @delete_tempfiles_on_error = !opts[:keep_tempfiles_on_error]
    end

    def data_only
      @data_only = true
      self
    end

    def parse(xl_file_path)
      raise 'Not .xlsx or .xlsm excel file' unless @skip_filename_check ||
                                              %w{.xlsx .xlsm}.include?(File.extname(xl_file_path))

      dir_path = File.join(File.dirname(xl_file_path), Dir::Tmpname.make_tmpname(['rubyXL', '.tmp'], nil))
      parse_success = false

      begin
        ::Zip::File.open(xl_file_path) { |zip_file|
          zip_file.each { |f|
            fpath = File.join(dir_path, f.name)
            FileUtils.mkdir_p(File.dirname(fpath))
            zip_file.extract(f, fpath) unless File.exist?(fpath)
          }
        }

        rels = RubyXL::RootRelationships.parse_file(dir_path)
        rels.load_related_files(dir_path)

        document_members = rels.related_files.values
        wb = document_members.find { |obj| obj.is_a?(RubyXL::Workbook) }
        wb.root_relationship_container = rels
        wb.filepath = xl_file_path
        wb.content_types = RubyXL::ContentTypes.parse_file(dir_path)
        wb.core_properties     = document_members.find { |obj| obj.is_a?(RubyXL::CoreProperties) }
        wb.document_properties = document_members.find { |obj| obj.is_a?(RubyXL::DocumentProperties) }

        unless @data_only
          wb.external_links.load_dir(dir_path)
          wb.external_links_rels.load_dir(dir_path)
  #        wb.drawings.load_dir(dir_path)
  #        wb.drawings_rels.load_dir(dir_path)
  #        wb.charts.load_dir(dir_path)
  #        wb.chart_rels.load_dir(dir_path)
  #        wb.worksheet_rels.load_dir(dir_path)
  #        wb.chartsheet_rels.load_dir(dir_path)
          wb.macros.load_file(dir_path, 'vbaProject.bin')
          wb.thumbnail.load_file(dir_path, 'thumbnail.jpeg')

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

        parse_success = true
      ensure
        FileUtils.remove_entry_secure(dir_path) if parse_success || @delete_tempfiles_on_error
      end

      return wb
    end

  end
end
