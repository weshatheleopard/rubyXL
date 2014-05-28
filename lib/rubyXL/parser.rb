require 'nokogiri'
require 'tmpdir'
require 'zip'

module RubyXL

  class Parser
    def self.parse(file_path, opts = {})
      self.new(opts).parse(file_path)
    end

    # +:data_only+ allows only the sheet data to be parsed, so as to speed up parsing.
    # However, using this option will result in date-formatted cells being interpreted as numbers.
    def initialize(opts = {})
      @opts = opts
    end

    def parse(xl_file_path)
      root = RubyXL::WorkbookRoot.parse_file(xl_file_path, @opts)

      wb = root.workbook

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

      wb
    end

  end
end
