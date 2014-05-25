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

        root = RubyXL::WorkbookRoot.new
        root.load_relationships(dir_path)

        wb = root.workbook
        wb.root = root
        wb.content_types = RubyXL::ContentTypes.parse_file(dir_path)

        rels = root.relationship_container

        wb.filepath = xl_file_path

        if wb.stylesheet then
          #fills out count information for each font, fill, and border
          wb.cell_xfs.each { |style|

            id = style.font_id
            wb.fonts[id].count += 1 #unless id.nil?

            id = style.fill_id
            wb.fills[id].count += 1 #unless id.nil?

            id = style.border_id
            wb.borders[id].count += 1 #unless id.nil?
          }
        end

#          sheet_path = sheet_rel.target
#          sheet_path = File.join('xl', sheet_path) if sheet_path[0] != '/'

        parse_success = true
      ensure
        FileUtils.remove_entry_secure(dir_path) if parse_success || @delete_tempfiles_on_error
      end

      return wb
    end

  end
end
