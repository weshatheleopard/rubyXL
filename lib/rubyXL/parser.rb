require 'rubygems'
require 'nokogiri'
require 'zip/zip' #rubyzip
require File.expand_path(File.join(File.dirname(__FILE__),'Hash'))

module RubyXL

  class Parser
    attr_reader :data_only, :num_sheets
    @@parsed_column_hash ={}
    # converts cell string (such as "AA1") to matrix indices
    def Parser.convert_to_index(cell_string)
      index = [-1,-1]
      if(cell_string =~ /^([A-Z]+)(\d+)$/)

        one = $1
        row = $2.to_i - 1 #-1 for 0 indexing
        col = 0
        i = 0
        if @@parsed_column_hash[one].nil?
          two = one.reverse #because of 26^i calculation
          two.each_byte do |c|
            int_val = c - 64 #converts A to 1
            col += int_val * 26**(i)
            i=i+1
          end
          @@parsed_column_hash[one] = col
        else
          col = @@parsed_column_hash[one]
        end
        col -= 1 #zer0 index
        index[0] = row
        index[1] = col
      end
      return index
    end


    # data_only allows only the sheet data to be parsed, so as to speed up parsing
    # However, using this option will result in date-formatted cells being interpreted as numbers
    def Parser.parse(file_path, opts = {})

      # options handling
      @data_only = opts.is_a?(TrueClass)||!!opts[:data_only]
      skip_filename_check = !!opts[:skip_filename_check]

      files = Parser.decompress(file_path, skip_filename_check)
      wb = Parser.fill_workbook(file_path, files)

      if(files['sharedString'] != nil)
        wb.num_strings = Integer(files['sharedString'].css('sst').attribute('count').value())
        wb.size = Integer(files['sharedString'].css('sst').attribute('uniqueCount').value())

        files['sharedString'].css('si').each do |node|
          unless node.css('r').empty?
            text = node.css('r t').children.to_a.join
            node.children.remove
            node << "<t xml:space=\"preserve\">#{text}</t>"
          end
        end

        string_nodes = files['sharedString'].css('si t')
        wb.shared_strings = {}
        string_nodes.each_with_index do |node,i|
          string = node.children.to_s
          wb.shared_strings[i] = string
          wb.shared_strings[string] = i
        end
      end
      #styles are needed for formatting reasons as that is how dates are determined
        styles = files['styles'].css('cellXfs xf')
        style_hash = Hash.xml_node_to_hash(files['styles'].root)
        fill_styles(wb,style_hash)

        #will be nil if these files do not exist
        wb.external_links = files['externalLinks']
        wb.drawings = files['drawings']
        wb.printer_settings = files['printerSettings']
        wb.worksheet_rels = files['worksheetRels']
        wb.macros = files['vbaProject']

      #for each worksheet:
      #1. find the dimensions of the data matrix
      #2. Fill in the matrix with data from worksheet/shared_string files
      #3. Apply styles
      wb.worksheets.each_index do |i|
        Parser.fill_worksheet(wb,i,files,wb.shared_strings)
      end

      return wb
    end

    private

    #fills hashes for various styles
    def Parser.fill_styles(wb,style_hash)
      ###NUM FORMATS###
      if style_hash[:numFmts].nil?
        style_hash[:numFmts] = {:attributes => {:count => 0}, :numFmt => []}
      elsif style_hash[:numFmts][:attributes][:count]==1
        style_hash[:numFmts][:numFmt] = [style_hash[:numFmts][:numFmt]]
      end
      wb.num_fmts = style_hash[:numFmts]

      ###FONTS###
      wb.fonts = {}
      if style_hash[:fonts][:attributes][:count]==1
        style_hash[:fonts][:font] = [style_hash[:fonts][:font]]
      end

      style_hash[:fonts][:font].each_with_index do |f,i|
        wb.fonts[i.to_s] = {:font=>f,:count=>0}
      end

      ###FILLS###
      wb.fills = {}
      if style_hash[:fills][:attributes][:count]==1
        style_hash[:fills][:fill] = [style_hash[:fills][:fill]]
      end

      style_hash[:fills][:fill].each_with_index do |f,i|
        wb.fills[i.to_s] = {:fill=>f,:count=>0}
      end

      ###BORDERS###
      wb.borders = {}
      if style_hash[:borders][:attributes][:count] == 1
        style_hash[:borders][:border] = [style_hash[:borders][:border]]
      end

      style_hash[:borders][:border].each_with_index do |b,i|
        wb.borders[i.to_s] = {:border=>b, :count=>0}
      end

      wb.cell_style_xfs = style_hash[:cellStyleXfs]
      wb.cell_xfs = style_hash[:cellXfs]
      wb.cell_styles = style_hash[:cellStyles]

      wb.colors = style_hash[:colors]

      #fills out count information for each font, fill, and border
      if wb.cell_xfs[:xf].is_a?(::Hash)
        wb.cell_xfs[:xf] = [wb.cell_xfs[:xf]]
      end
      wb.cell_xfs[:xf].each do |style|
        id = style[:attributes][:fontId].to_s
        unless id.nil?
          wb.fonts[id][:count] += 1
        end

        id = style[:attributes][:fillId].to_s
        unless id.nil?
          wb.fills[id][:count] += 1
        end

        id = style[:attributes][:borderId].to_s
        unless id.nil?
          wb.borders[id][:count] += 1
        end
      end
    end

    # i is the sheet number
    # files is the hash which includes information for each worksheet
    # shared_strings has group of indexed strings which the cells reference
    def Parser.fill_worksheet(wb,i,files,shared_strings)
      wb.worksheets[i] = Parser.create_matrix(wb, i, files)
      j = i+1

      namespaces = files[j].root.namespaces()
      unless @data_only
        sheet_views_node= files[j].xpath('/xmlns:worksheet/xmlns:sheetViews[xmlns:sheetView]',namespaces).first
        wb.worksheets[i].sheet_view = Hash.xml_node_to_hash(sheet_views_node)[:sheetView]

        ##col styles##
        cols_node_set = files[j].xpath('/xmlns:worksheet/xmlns:cols',namespaces)
        unless cols_node_set.empty?
          wb.worksheets[i].cols= Hash.xml_node_to_hash(cols_node_set.first)[:col]
        end
        ##end col styles##

        ##merge_cells##
        merge_cells_node = files[j].xpath('/xmlns:worksheet/xmlns:mergeCells[xmlns:mergeCell]',namespaces)
        unless merge_cells_node.empty?
          wb.worksheets[i].merged_cells = Hash.xml_node_to_hash(merge_cells_node.first)[:mergeCell]
        end
        ##end merge_cells##

        ##sheet_view pane##
        pane_data = wb.worksheets[i].sheet_view[:pane]
        wb.worksheets[i].pane = pane_data
        ##end sheet_view pane##

        ##data_validation##
        data_validations_node = files[j].xpath('/xmlns:worksheet/xmlns:dataValidations[xmlns:dataValidation]',namespaces)
        unless data_validations_node.empty?
          wb.worksheets[i].validations = Hash.xml_node_to_hash(data_validations_node.first)[:dataValidation]
        else
          wb.worksheets[i].validations=nil
        end
        ##end data_validation##

        #extLst
        ext_list_node=files[j].xpath('/xmlns:worksheet/xmlns:extLst',namespaces)
        unless ext_list_node.empty?
          wb.worksheets[i].extLst = Hash.xml_node_to_hash(ext_list_node.first)
        else
          wb.worksheets[i].extLst=nil
        end
        #extLst

        ##legacy drawing##
        legacy_drawing_node = files[j].xpath('/xmlns:worksheet/xmlns:legacyDrawing',namespaces)
        unless legacy_drawing_node.empty?
          wb.worksheets[i].legacy_drawing = Hash.xml_node_to_hash(legacy_drawing_node.first)
        else
          wb.worksheets[i].legacy_drawing = nil
        end
        ##end legacy drawing
      end


      row_data = files[j].xpath('/xmlns:worksheet/xmlns:sheetData/xmlns:row[xmlns:c[xmlns:v]]',namespaces)
      row_data.each do |row|
        unless @data_only
          ##row styles##
          row_style = '0'
          row_attributes = row.attributes
          unless row_attributes['s'].nil?
            row_style = row_attributes['s'].value
          end

          wb.worksheets[i].row_styles[row_attributes['r'].content] = { :style => row_style  }

          if !row_attributes['ht'].nil?  && (!row_attributes['ht'].content.nil? || row_attributes['ht'].content.strip != "" )
            wb.worksheets[i].change_row_height(Integer(row_attributes['r'].content)-1,
              Float(row_attributes['ht'].content))
          end
          ##end row styles##
        end

        unless @data_only
          c_row = row.search('./xmlns:c')
        else
          c_row = row.search('./xmlns:c[xmlns:v[text()]]')
        end
        c_row.each do |value|
          #attributes is from the excel cell(c) and is basically location information and style and type
          value_attributes= value.attributes
          # r attribute contains the location like A1
          cell_index = Parser.convert_to_index(value_attributes['r'].content)
          style_index = 0

          # t is optional and contains the type of the cell
          data_type = value_attributes['t'].content if value_attributes['t']
          element_hash ={}
          value.children.each do |node|
            element_hash["#{node.name()}_element"]=node
          end
          # v is the value element that is part of the cell
          if element_hash["v_element"]
            v_element_content = element_hash["v_element"].content
          else
            v_element_content=""
          end
          if v_element_content =="" #no data
            cell_data = nil
          elsif data_type == 's' #shared string
            str_index = Integer(v_element_content)
            cell_data = shared_strings[str_index].to_s
          elsif data_type=='str' #raw string
            cell_data = v_element_content
          elsif data_type=='e' #error
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

          wb.worksheets[i].sheet_data[cell_index[0]][cell_index[1]] =
            Cell.new(wb.worksheets[i],cell_index[0],cell_index[1],cell_data,cell_formula,
              data_type,style_index,cell_formula_attr)
          cell = wb.worksheets[i].sheet_data[cell_index[0]][cell_index[1]]
        end
      end
    end

    def Parser.decompress(file_path, skip_filename_check = false)
      #ensures it is an xlsx/xlsm file
      if(file_path =~ /(.+)\.xls(x|m)/)
        dir_path = $1.to_s
      else
        if skip_filename_check
          dir_path = file_path
        else
          raise 'Not .xlsx or .xlsm excel file'
        end
      end

      dir_path = File.join(File.dirname(dir_path), make_safe_name(Time.now.to_s))
      #copies excel file to zip file in same directory
      zip_path = dir_path + '.zip'

      FileUtils.cp(file_path,zip_path)

      MyZip.new.unzip(zip_path,dir_path)
      File.delete(zip_path)

      files = Hash.new

      files['app'] = Nokogiri::XML.parse(File.open(File.join(dir_path,'docProps','app.xml'),'r'))
      files['core'] = Nokogiri::XML.parse(File.open(File.join(dir_path,'docProps','core.xml'),'r'))

      files['workbook'] = Nokogiri::XML.parse(File.open(File.join(dir_path,'xl','workbook.xml'),'r'))

      if(File.exist?(File.join(dir_path,'xl','sharedStrings.xml')))
        files['sharedString'] = Nokogiri::XML.parse(File.open(File.join(dir_path,'xl','sharedStrings.xml'),'r'))
      end

      unless @data_only
        #preserves external links
        if File.directory?(File.join(dir_path,'xl','externalLinks'))
          files['externalLinks'] = {}
          ext_links_path = File.join(dir_path,'xl','externalLinks')
          FileUtils.mkdir_p(ext_links_path)
          files['externalLinks']['rels'] = []
          dir = Dir.new(ext_links_path).entries.reject {|f| [".", "..", ".DS_Store", "_rels"].include? f}

          dir.each_with_index do |link,i|
            files['externalLinks'][i+1] = File.read(File.join(ext_links_path,link))
          end

          if File.directory?(File.join(ext_links_path,'_rels'))
            dir = Dir.new(File.join(ext_links_path,'_rels')).entries.reject{|f| [".","..",".DS_Store"].include? f}
            dir.each_with_index do |rel,i|
              files['externalLinks']['rels'][i+1] = File.read(File.join(ext_links_path,'_rels',rel))
            end
          end
        end

        if File.directory?(File.join(dir_path,'xl','drawings'))
          files['drawings'] = {}
          drawings_path = File.join(dir_path,'xl','drawings','_rels')
          FileUtils.mkdir_p(drawings_path)
          dir = Dir.new(drawings_path).entries.reject {|f| [".", "..", ".DS_Store"].include? f}
          dir.each_with_index do |draw,i|
            files['drawings'][i+1] = File.read(File.join(drawings_path,draw))
          end
        end

        if File.directory?(File.join(dir_path,'xl','printerSettings'))
          files['printerSettings'] = {}
          printer_path = File.join(dir_path,'xl','printerSettings')
          FileUtils.mkdir_p(printer_path)
          dir = Dir.new(printer_path).entries.reject {|f| [".","..",".DS_Store"].include? f}

          dir.each_with_index do |print, i|
            files['printerSettings'][i+1] = File.open(File.join(printer_path,print), 'rb').read
          end
        end

        if File.directory?(File.join(dir_path,"xl",'worksheets','_rels'))
          files['worksheetRels'] = {}
          worksheet_rels_path = File.join(dir_path,'xl','worksheets','_rels')
          FileUtils.mkdir_p(worksheet_rels_path)
          dir = Dir.new(worksheet_rels_path).entries.reject {|f| [".","..",".DS_Store"].include? f}
          dir.each_with_index do |rel, i|
            files['worksheetRels'][i+1] = File.read(File.join(worksheet_rels_path,rel))
          end
        end

        if File.exist?(File.join(dir_path,'xl','vbaProject.bin'))
          files['vbaProject'] = File.open(File.join(dir_path,"xl","vbaProject.bin"),'rb').read
        end
      end
      files['styles'] = Nokogiri::XML.parse(File.open(File.join(dir_path,'xl','styles.xml'),'r'))
      @num_sheets = files['workbook'].css('sheets').children.size
      @num_sheets = Integer(@num_sheets)

      #adds all worksheet xml files to files hash
      i=1
      1.upto(@num_sheets) do
        filename = 'sheet'+i.to_s
        files[i] = Nokogiri::XML.parse(File.open(File.join(dir_path,'xl','worksheets',filename+'.xml'),'r'))
        i=i+1
      end

      FileUtils.rm_rf(dir_path)

      return files
    end

    def Parser.fill_workbook(file_path, files)
      wb = Workbook.new([nil],file_path)

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
      wb.defined_names = files['workbook'].css('definedNames').to_s
      wb.date1904 = files['workbook'].css('workbookPr').attribute('date1904').to_s == '1'

      wb.worksheets = Array.new(@num_sheets) #array of Worksheet objs
      wb
    end

    #sheet_names, dimensions
    def Parser.create_matrix(wb,i, files)
      sheet_names = files['app'].css('TitlesOfParts vt|vector vt|lpstr').children
      sheet = Worksheet.new(wb,sheet_names[i].to_s,[])

      dimensions = files[i+1].css('dimension').attribute('ref').to_s
      if(dimensions =~ /^([A-Z]+\d+:)?([A-Z]+\d+)$/)
        index = convert_to_index($2)

        rows = index[0]+1
        cols = index[1]+1

        #creates matrix filled with nils
        rows.times {sheet.sheet_data << Array.new(cols)}
      else
        raise 'invalid file'
      end
      sheet
    end

    def Parser.safe_filename(name, allow_mb_chars=false)
      # "\w" represents [0-9A-Za-z_] plus any multi-byte char
      regexp = allow_mb_chars ? /[^\w]/ : /[^0-9a-zA-Z\_]/
      name.gsub(regexp, "_")
    end

    # Turns the passed in string into something safe for a filename
    def Parser.make_safe_name(name, allow_mb_chars=false)
      ext = safe_filename(File.extname(name), allow_mb_chars).gsub(/^_/, '.')
      "#{safe_filename(name.gsub(ext, ""), allow_mb_chars)}#{ext}".gsub(/\(/, '_').gsub(/\)/, '_').gsub(/__+/, '_').gsub(/^_/, '').gsub(/_$/, '')
    end

  end
end
