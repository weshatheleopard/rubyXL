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

      FileUtils.rm_rf(dir_path)

      return wb
    end

    private

    # Parse the incoming +worksheet_xml+ into a new +Worksheet+ object 
    def parse_worksheet(wb, i, worksheet_xml, worksheet_name, sheet_id)
      worksheet = Worksheet.new(wb, worksheet_name)
      wb.worksheets[i] = worksheet # Due to "validate_workbook" issues. Should remove that validation eventually.
      worksheet.sheet_id = sheet_id
      dimensions = RubyXL::Reference.new(worksheet_xml.css('dimension').attribute('ref').value)
      cols = dimensions.last_col

      # Create empty arrays for workcells. Using +downto()+ here so memory for +sheet_data[]+ is
      # allocated on the first iteration (in case of +upto()+, +sheet_data[]+ would end up being
      # reallocated on every iteration).
      dimensions.last_row.downto(0) { |i| worksheet.sheet_data[i] = Array.new(cols + 1) }

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

      worksheet_xml.xpath(row_xpath, namespaces).each { |row|
        unless @data_only
          ##row styles##
          row_attributes = row.attributes
          row_style = row_attributes['s'] && row_attributes['s'].value || '0'

          worksheet.row_styles[row_attributes['r'].content] = { :style => row_style  }

          if !row_attributes['ht'].nil?  && (!row_attributes['ht'].content.nil? || row_attributes['ht'].content.strip != "" )
            worksheet.change_row_height(Integer(row_attributes['r'].value) - 1, Float(row_attributes['ht'].value))
          end
          ##end row styles##
        end

        row.search(cell_xpath).each { |value|
          #attributes is from the excel cell(c) and is basically location information and style and type
          value_attributes = value.attributes
          # r attribute contains the location like A1
          cell_index = RubyXL::Reference.ref2ind(value_attributes['r'].content)
          style_index = 0
          # t is optional and contains the type of the cell
          data_type = value_attributes['t'].content if value_attributes['t']
          element_hash ={}

          value.children.each { |node| element_hash["#{node.name()}_element"] = node }

          # v is the value element that is part of the cell
          if element_hash["v_element"]
            v_element_content = element_hash["v_element"].content
          else
            v_element_content=""
          end

          if v_element_content == "" # no data
            cell_data = nil
          elsif data_type == RubyXL::Cell::SHARED_STRING
            str_index = Integer(v_element_content)
            cell_data = wb.shared_strings[str_index].to_s
          elsif data_type == RubyXL::Cell::RAW_STRING
            cell_data = v_element_content
          elsif data_type == RubyXL::Cell::ERROR
            cell_data = v_element_content
          else# (value.css('v').to_s != "") && (value.css('v').children.to_s != "") #is number
            data_type = ''
            if(v_element_content =~ /\./ or v_element_content =~ /\d+e\-?\d+/i) #is float
              cell_data = Float(v_element_content)
            else
              cell_data = Integer(v_element_content)
            end
          end

          # f is the formula element
          cell_formula = nil
          fmla_css = element_hash["f_element"]
          if fmla_css && fmla_css.content
            fmla_css_content= fmla_css.content
            if(fmla_css_content != "")
              cell_formula = fmla_css_content
              cell_formula_attr = {}
              fmla_css_attributes = fmla_css.attributes
              cell_formula_attr['t'] = fmla_css_attributes['t'].content if fmla_css_attributes['t']
              cell_formula_attr['ref'] = fmla_css_attributes['ref'].content if fmla_css_attributes['ref']
              cell_formula_attr['si'] = fmla_css_attributes['si'].content if fmla_css_attributes['si']
            end
          end

          style_index = value['s'].to_i #nil goes to 0 (default)

          worksheet.sheet_data[cell_index[0]][cell_index[1]] =
            Cell.new(worksheet,cell_index[0],cell_index[1],cell_data,cell_formula,
              data_type,style_index,cell_formula_attr)
          cell = worksheet.sheet_data[cell_index[0]][cell_index[1]]
        }
      }

      worksheet
    end

    def self.attr_int(node, attr_name) 
      attr = node.attributes[attr_name]
      attr && Integer(attr.value)
    end

    def self.attr_float(node, attr_name) 
      attr = node.attributes[attr_name]
      attr && Float(attr.value)
    end

    def self.attr_string(node, attr_name) 
      attr = node.attributes[attr_name]
      attr && attr.value
    end

  end
end

