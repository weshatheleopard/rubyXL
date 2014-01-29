require 'rubygems'
require 'nokogiri'
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

    def parse(file_path, opts = {})
      raise 'Not .xlsx or .xlsm excel file' unless @skip_filename_check ||
                                              %w{.xlsx .xlsm}.include?(File.extname(file_path))

      dir_path = File.join(File.dirname(file_path), Dir::Tmpname.make_tmpname(['rubyXL', '.tmp'], nil))

      MyZip.new.unzip(file_path, dir_path)

      workbook_file = Nokogiri::XML.parse(File.open(File.join(dir_path, 'xl', 'workbook.xml'), 'r'))

      wb = RubyXL::Workbook.parse(workbook_file)
      wb.filepath = file_path

      rels_doc = Nokogiri::XML.parse(File.open(File.join(dir_path, 'xl', '_rels', 'workbook.xml.rels'), 'r'))

      if(File.exist?(File.join(dir_path,'xl','sharedStrings.xml')))
        shared_string_file = Nokogiri::XML.parse(File.open(File.join(dir_path,'xl','sharedStrings.xml'),'r'))
      end

      unless @data_only
        wb.media = RubyXL::GenericStorage.new(File.join('xl', 'media')).binary.load_dir(dir_path)
        wb.external_links = RubyXL::GenericStorage.new(File.join('xl', 'externalLinks')).load_dir(dir_path)
        wb.external_links_rels = RubyXL::GenericStorage.new(File.join('xl', 'externalLinks', '_rels')).load_dir(dir_path)
        wb.drawings = RubyXL::GenericStorage.new(File.join('xl', 'drawings')).load_dir(dir_path)
        wb.drawings_rels = RubyXL::GenericStorage.new(File.join('xl', 'drawings', '_rels')).load_dir(dir_path)
        wb.charts = RubyXL::GenericStorage.new(File.join('xl', 'charts')).load_dir(dir_path)
        wb.chart_rels = RubyXL::GenericStorage.new(File.join('xl', 'charts', '_rels')).load_dir(dir_path)
        wb.printer_settings = RubyXL::GenericStorage.new(File.join('xl', 'printerSettings')).binary.load_dir(dir_path)
        wb.worksheet_rels = RubyXL::GenericStorage.new(File.join('xl', 'worksheets', '_rels')).load_dir(dir_path)
        wb.macros = RubyXL::GenericStorage.new('xl').binary.load_file(dir_path, 'vbaProject.bin')
        wb.theme = RubyXL::GenericStorage.new(File.join('xl', 'theme')).load_file(dir_path, 'theme1.xml')

        core_file = Nokogiri::XML.parse(File.open(File.join(dir_path, 'docProps', 'core.xml'), 'r'))
        wb.creator = core_file.css('dc|creator').children.to_s
        wb.modifier = core_file.css('cp|last_modified_by').children.to_s
        wb.created_at = core_file.css('dcterms|created').children.to_s
        wb.modified_at = core_file.css('dcterms|modified').children.to_s

        app_file = Nokogiri::XML.parse(File.open(File.join(dir_path, 'docProps', 'app.xml'), 'r'))
        wb.company = app_file.css('Company').children.to_s
        wb.application = app_file.css('Application').children.to_s
        wb.appversion = app_file.css('AppVersion').children.to_s
      end

      styles_xml = Nokogiri::XML.parse(File.open(File.join(dir_path, 'xl', 'styles.xml'), 'r'))

      wb.date1904 = workbook_file.css('workbookPr').attribute('date1904').to_s == '1'

      wb.shared_strings_XML = shared_string_file.to_s

      unless shared_string_file.nil?
        sst = shared_string_file.css('sst')

        # According to http://msdn.microsoft.com/en-us/library/office/gg278314.aspx,
        # these attributes may be either both missing, or both present. Need to validate.
        wb.shared_strings.count_attr = sst.attribute('count').value.to_i
        wb.shared_strings.unique_count_attr = sst.attribute('uniqueCount').value.to_i

        # Note that the strings may contain text formatting, such as changing font color/properties
        # in the middle of the string. We do not support that in this gem... at least yet!
        # If you save the file, this formatting will be destoyed.
        shared_string_file.css('si').each_with_index { |node, i|
          wb.shared_strings.add(node.css('t').inject(''){ |s, c| s + c.text }, i)
        }

      end

      wb.stylesheet = RubyXL::Stylesheet.parse(styles_xml.root)

      #fills out count information for each font, fill, and border
      wb.cell_xfs.each { |style|
        id = style.font_id
        wb.fonts[id].count += 1 #unless id.nil?

        id = style.fill_id
        wb.fills[id].count += 1 #unless id.nil?

        id = style.border_id
        wb.borders[id].count += 1 #unless id.nil?
      }

      # Not sure why they were getting sheet names from god knows where.
      # There *may* have been a good reason behind it, so not tossing this code out entirely yet.
      # sheet_names = app_file.css('TitlesOfParts vt|vector vt|lpstr').children

      workbook_file.css('sheets sheet').each_with_index { |sheet_node, i|
        sheet_rid = sheet_node.attributes['id'].value 
        sheet_file_path = rels_doc.css("Relationships Relationship[Id=#{sheet_rid}]").first.attributes['Target']
        worksheet_xml = Nokogiri::XML.parse(File.read(File.join(dir_path, 'xl', sheet_file_path)))

        worksheet = RubyXL::Worksheet.parse(worksheet_xml.root)
        worksheet.sheet_data.rows.each { |r| r && r.cells.each { |c| c.worksheet = worksheet unless c.nil? } }
        wb.worksheets[i] = worksheet
        worksheet.workbook = wb
        worksheet.sheet_name = sheet_node.attributes['name'].value
        worksheet.sheet_id = sheet_node.attributes['sheetId'].value
      }

      FileUtils.remove_entry_secure(dir_path)

      return wb
    end

  end
end

