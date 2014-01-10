require 'rubygems'
require 'nokogiri'
require 'zip'
require 'rubyXL/hash'
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

      files = decompress(file_path)
      wb = Workbook.new([], file_path)
      fill_workbook(wb, files)

      shared_string_file = files['sharedString']
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

      fills = files['styles'].css('fills fill')
      wb.fills = fills.collect { |node| RubyXL::Fill.parse(node) }

      colors = files['styles'].css('colors').first

      if colors then
        colors.element_children.each { |color_type_node|
          wb.colors[color_type_node.name] ||= []
          color_type_node.element_children.each { |color_node|
            wb.colors[color_type_node.name] << RubyXL::Color.parse(color_node)
          }
        }
      end

      borders = files['styles'].css('borders border')
      wb.borders = borders.collect { |node| RubyXL::Border.parse(node) }

      fonts = files['styles'].css('fonts font')
      wb.fonts = fonts.collect { |node| RubyXL::Font.parse(node) }

      cell_styles = files['styles'].css('cellStyles cellStyle')
      wb.cell_styles = cell_styles.collect { |node| RubyXL::CellStyle.parse(node) }

      num_fmts = files['styles'].css('numFmts numFmt')
      wb.num_fmts = num_fmts.collect { |node| RubyXL::NumFmt.parse(node) }

      fill_styles(wb, Hash.xml_node_to_hash(files['styles'].root))

      wb.media = files['media']
      wb.external_links = files['externalLinks']
      wb.external_links_rels = files['externalLinksRels']
      wb.drawings = files['drawings']
      wb.drawings_rels = files['drawingsRels']
      wb.charts = files['charts']
      wb.chart_rels = files['chartRels']
      wb.printer_settings = files['printerSettings']
      wb.worksheet_rels = files['worksheetRels']
      wb.macros = files['vbaProject']
      wb.theme = files['theme']

      # Not sure why they were getting sheet names from god knows where.
      # There *may* have been a good reason behind it, so not tossing this code out entirely yet.
      # sheet_names = files['app'].css('TitlesOfParts vt|vector vt|lpstr').children

      files['workbook'].css('sheets sheet').each_with_index { |sheet_node, i|
        parse_worksheet(wb, i, files['worksheets'][i], sheet_node.attributes['name'].value,
                               sheet_node.attributes['sheetId'].value )
      }

      return wb
    end

    private

    #fills hashes for various styles
    def fill_styles(wb,style_hash)
      wb.cell_style_xfs = style_hash[:cellStyleXfs]
      wb.cell_xfs = style_hash[:cellXfs]

      #fills out count information for each font, fill, and border
      if wb.cell_xfs[:xf].is_a?(::Hash)
        wb.cell_xfs[:xf] = [wb.cell_xfs[:xf]]
      end

      wb.cell_xfs[:xf].each do |style|
        id = Integer(style[:attributes][:fontId])
        wb.fonts[id].count += 1 unless id.nil?

        id = style[:attributes][:fillId]
        wb.fills[id].count += 1 unless id.nil?

        id = style[:attributes][:borderId]
        wb.borders[id].count += 1 unless id.nil?
      end

    end

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

        #extLst
        ext_list_node = worksheet_xml.xpath('/xmlns:worksheet/xmlns:extLst', namespaces)
        unless ext_list_node.empty?
          worksheet.extLst = Hash.xml_node_to_hash(ext_list_node.first)
        else
          worksheet.extLst = nil
        end
        #extLst

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

    def decompress(file_path)
      dir_path = file_path

      #ensures it is an xlsx/xlsm file
      if (file_path =~ /(.+)\.xls(x|m)/) then
        dir_path = $1.to_s
      else
        raise 'Not .xlsx or .xlsm excel file' unless @skip_filename_check
      end

      dir_path = File.join(File.dirname(dir_path), Dir::Tmpname.make_tmpname(['rubyXL', '.tmp'], nil))

      #copies excel file to zip file in same directory
      zip_path = dir_path + '.zip'

      FileUtils.cp(file_path,zip_path)

      MyZip.new.unzip(zip_path,dir_path)
      File.delete(zip_path)

      files = Hash.new

      files['app'] = Nokogiri::XML.parse(File.open(File.join(dir_path,'docProps','app.xml'),'r'))
      files['core'] = Nokogiri::XML.parse(File.open(File.join(dir_path,'docProps','core.xml'),'r'))
      files['workbook'] = Nokogiri::XML.parse(File.open(File.join(dir_path,'xl','workbook.xml'),'r'))
      files['workbook_rels'] = Nokogiri::XML.parse(File.open(File.join(dir_path, 'xl', '_rels', 'workbook.xml.rels'), 'r'))

      if(File.exist?(File.join(dir_path,'xl','sharedStrings.xml')))
        files['sharedString'] = Nokogiri::XML.parse(File.open(File.join(dir_path,'xl','sharedStrings.xml'),'r'))
      end

      unless @data_only
        files['media'] = RubyXL::GenericStorage.new(File.join('xl', 'media')).binary.load_dir(dir_path)
        files['externalLinks'] = RubyXL::GenericStorage.new(File.join('xl', 'externalLinks')).load_dir(dir_path)
        files['externalLinksRels'] = RubyXL::GenericStorage.new(File.join('xl', 'externalLinks', '_rels')).load_dir(dir_path)
        files['drawings'] = RubyXL::GenericStorage.new(File.join('xl', 'drawings')).load_dir(dir_path)
        files['drawingsRels'] = RubyXL::GenericStorage.new(File.join('xl', 'drawings', '_rels')).load_dir(dir_path)
        files['charts'] = RubyXL::GenericStorage.new(File.join('xl', 'charts')).load_dir(dir_path)
        files['chartRels'] = RubyXL::GenericStorage.new(File.join('xl', 'charts', '_rels')).load_dir(dir_path)
        files['printerSettings'] = RubyXL::GenericStorage.new(File.join('xl', 'printerSettings')).binary.load_dir(dir_path)
        files['worksheetRels'] = RubyXL::GenericStorage.new(File.join('xl', 'worksheets', '_rels')).load_dir(dir_path)
        files['vbaProject'] = RubyXL::GenericStorage.new('xl').binary.load_file(dir_path, 'vbaProject.bin')
        files['theme'] = RubyXL::GenericStorage.new(File.join('xl', 'theme')).load_file(dir_path, 'theme1.xml')
      end

      files['styles'] = Nokogiri::XML.parse(File.open(File.join(dir_path,'xl','styles.xml'),'r'))

      files['worksheets'] = []
      rels_doc = files['workbook_rels']

      files['workbook'].css('sheets sheet').each_with_index { |sheet, ind|
        sheet_rid = sheet.attributes['id'].value 
        sheet_file_path = rels_doc.css("Relationships Relationship[Id=#{sheet_rid}]").first.attributes['Target']
        files['worksheets'][ind] = Nokogiri::XML.parse(File.read(File.join(dir_path, 'xl', sheet_file_path)))
      }

      FileUtils.rm_rf(dir_path)

      return files
    end

    def fill_workbook(wb, files)
      unless @data_only
        wb.creator = files['core'].css('dc|creator').children.to_s
        wb.modifier = files['core'].css('cp|last_modified_by').children.to_s
        wb.created_at = files['core'].css('dcterms|created').children.to_s
        wb.modified_at = files['core'].css('dcterms|modified').children.to_s

        wb.company = files['app'].css('Company').children.to_s
        wb.application = files['app'].css('Application').children.to_s
        wb.appversion = files['app'].css('AppVersion').children.to_s
      end

      wb.shared_strings_XML = files['sharedString'].to_s

      defined_names = files['workbook'].css('definedNames definedName')
      wb.defined_names = defined_names.collect { |node| RubyXL::DefinedName.parse(node) }

      wb.date1904 = files['workbook'].css('workbookPr').attribute('date1904').to_s == '1'
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

