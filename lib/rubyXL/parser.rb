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

      # options handling

      wb = Workbook.new([], file_path)

      raise 'Not .xlsx or .xlsm excel file' unless @skip_filename_check ||
                                              %w{.xlsx .xlsm}.include?(File.extname(file_path))

      dir_path = File.join(File.dirname(file_path), Dir::Tmpname.make_tmpname(['rubyXL', '.tmp'], nil))

      MyZip.new.unzip(file_path, dir_path)

      files = {}

      workbook_file = Nokogiri::XML.parse(File.open(File.join(dir_path,'xl','workbook.xml'),'r'))
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

      defined_names = workbook_file.css('definedNames definedName')
      wb.defined_names = defined_names.collect { |node| RubyXL::DefinedName.parse(node) }

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

      fills = styles_xml.css('fills fill')
      wb.fills = fills.collect { |node| RubyXL::Fill.parse(node) }

      colors = styles_xml.css('colors').first

      if colors then
        colors.element_children.each { |color_type_node|
          wb.colors[color_type_node.name] ||= []
          color_type_node.element_children.each { |color_node|
            wb.colors[color_type_node.name] << RubyXL::Color.parse(color_node)
          }
        }
      end

      borders = styles_xml.css('borders border')
      wb.borders = borders.collect { |node| RubyXL::Border.parse(node) }

      fonts = styles_xml.css('fonts font')
      wb.fonts = fonts.collect { |node| RubyXL::Font.parse(node) }

      cell_styles = styles_xml.css('cellStyles cellStyle')
      wb.cell_styles = cell_styles.collect { |node| RubyXL::CellStyle.parse(node) }

      num_fmts = styles_xml.css('numFmts numFmt')
      wb.num_fmts = num_fmts.collect { |node| RubyXL::NumFmt.parse(node) }

      csxfs = styles_xml.css('cellStyleXfs xf')
      wb.cell_style_xfs = csxfs.collect { |node| RubyXL::XF.parse(node) }

      cxfs = styles_xml.css('cellXfs xf')
      wb.cell_xfs = cxfs.collect { |node| RubyXL::XF.parse(node) }

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
        parse_worksheet(wb, i, worksheet_xml, sheet_node.attributes['name'].value,
                               sheet_node.attributes['sheetId'].value )
      }

      FileUtils.remove_entry_secure(dir_path)

      return wb
    end

    private

    # Parse the incoming +worksheet_xml+ into a new +Worksheet+ object 
    def parse_worksheet(wb, i, worksheet_xml, worksheet_name, sheet_id)
      worksheet = Worksheet.new(:workbook => wb, :name => worksheet_name)
      wb.worksheets[i] = worksheet # Due to "validate_workbook" issues. Should remove that validation eventually.
      worksheet.sheet_id = sheet_id

      dimensions_node = worksheet_xml.css('dimension')
      return nil if dimensions_node.empty? # Temporary plug for Issue #76

# Technically, we don't even need dimensions anymore since we are not pre-creating the array.
#      dimensions = RubyXL::Reference.new(dimensions_node.attribute('ref').value)

      namespaces = worksheet_xml.root.namespaces

      if @data_only
        row_xpath = '/xmlns:worksheet/xmlns:sheetData/xmlns:row[xmlns:c[xmlns:v]]'
        cell_xpath = './xmlns:c[xmlns:v[text()]]'
      else
        row_xpath = '/xmlns:worksheet/xmlns:sheetData/xmlns:row'
        cell_xpath = './xmlns:c'

        sheet_views_nodes = worksheet_xml.xpath('/xmlns:worksheet/xmlns:sheetViews/xmlns:sheetView', namespaces)
        worksheet.sheet_views = sheet_views_nodes.collect { |node| RubyXL::SheetView.parse(node) }

        col_node_set = worksheet_xml.xpath('/xmlns:worksheet/xmlns:cols/xmlns:col',namespaces)
        worksheet.column_ranges = col_node_set.collect { |col_node| RubyXL::ColumnRange.parse(col_node) }

        merged_cells_nodeset = worksheet_xml.xpath('/xmlns:worksheet/xmlns:mergeCells/xmlns:mergeCell', namespaces)
        worksheet.merged_cells = merged_cells_nodeset.collect { |child| RubyXL::Reference.new(child.attributes['ref'].text) }

#        worksheet.pane = worksheet.sheet_view[:pane]

        data_validations = worksheet_xml.xpath('/xmlns:worksheet/xmlns:dataValidations/xmlns:dataValidation', namespaces)
        worksheet.validations = data_validations.collect { |node| RubyXL::DataValidation.parse(node) }

# Currently  not working #TODO#
#        ext_list_node = worksheet_xml.xpath('/xmlns:worksheet/xmlns:extLst', namespaces)

        legacy_drawing_nodes = worksheet_xml.xpath('/xmlns:worksheet/xmlns:legacyDrawing', namespaces)
        worksheet.legacy_drawings = legacy_drawing_nodes.collect { |node| RubyXL::LegacyDrawing.parse(node) }

        drawing_nodes = worksheet_xml.xpath('/xmlns:worksheet/xmlns:drawing', namespaces)
        worksheet.drawings = drawing_nodes.collect { |n| n.attributes['id'] }

      end

      sheet_data = worksheet_xml.xpath('/xmlns:worksheet/xmlns:sheetData', namespaces)
      worksheet.sheet_data = RubyXL::SheetData.parse(sheet_data.first)

      test = RubyXL::Worksheet.parse(worksheet_xml.root)
      test.workbook = wb

      worksheet
    end

  end
end

