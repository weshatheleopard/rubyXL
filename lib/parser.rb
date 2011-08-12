require 'rubygems'
require 'nokogiri'
require 'zip/zip' #rubyzip
require File.expand_path(File.join(File.dirname(__FILE__),'Hash'))

module RubyXL

  class Parser
    attr_reader :data_only, :num_sheets

    # converts cell string (such as "AA1") to matrix indices
    def Parser.convert_to_index(cell_string)
      index = Array.new(2)
      index[0]=-1
      index[1]=-1
      if(cell_string =~ /^([A-Z]+)(\d+)$/)
        one = $1.to_s
        row = Integer($2) - 1 #-1 for 0 indexing
        col = 0
        i = 0
        one = one.reverse #because of 26^i calculation
        one.each_byte do |c|
          int_val = c - 64 #converts A to 1
          col += int_val * 26**(i)
          i=i+1
        end
        col -= 1 #zer0 index
        index[0] = row
        index[1] = col
      end
      return index
    end


    # data_only allows only the sheet data to be parsed, so as to speed up parsing
    def Parser.parse(file_path, data_only=false)
      @data_only = data_only
      files = Parser.decompress(file_path)
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
        

        shared_strings = files['sharedString'].css('si t').children
        wb.shared_strings = {}
        shared_strings.each_with_index do |string,i|
          wb.shared_strings[i] = string.to_s
          wb.shared_strings[string.to_s] = i
        end
      end

      unless @data_only
        styles = files['styles'].css('cellXfs xf')
        style_hash = Hash.xml_node_to_hash(files['styles'].root)
        fill_styles(wb,style_hash)

        #will be nil if these files do not exist
        wb.external_links = files['externalLinks']
        wb.drawings = files['drawings']
        wb.printer_settings = files['printerSettings']
        wb.worksheet_rels = files['worksheetRels']
        wb.macros = files['vbaProject']
      end

      #for each worksheet:
      #1. find the dimensions of the data matrix
      #2. Fill in the matrix with data from worksheet/shared_string files
      #3. Apply styles
      wb.worksheets.each_index do |i|
        Parser.fill_worksheet(wb,i,files,shared_strings)
      end

      return wb
    end

    private

    #fills hashes for various styles
    def Parser.fill_styles(wb,style_hash)
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

      unless @data_only
        hash = Hash.xml_node_to_hash(files[j].root)

        wb.worksheets[i].sheet_view = hash[:sheetViews][:sheetView]

        ##col styles##
        col_data = hash[:cols]
        unless col_data.nil?
          wb.worksheets[i].cols=col_data[:col]
        end
        ##end col styles##

        ##merge_cells##
        merge_data = hash[:mergeCells]
        unless merge_data.nil?
          wb.worksheets[i].merged_cells = merge_data[:mergeCell]
        end
        ##end merge_cells##

        ##sheet_view pane##
        pane_data = hash[:sheetViews][:sheetView][:pane]
        wb.worksheets[i].pane = pane_data
        ##end sheet_view pane##

        ##data_validation##
        data_validation = hash[:dataValidations]
        unless data_validation.nil?
          data_validation = data_validation[:dataValidation]
        end
        wb.worksheets[i].validations = data_validation
        ##end data_validation##

        #extLst
        wb.worksheets[i].extLst = hash[:extLst]
        #extLst

        ##legacy drawing##
        drawing = hash[:legacyDrawing]
        wb.worksheets[i].legacy_drawing = drawing
        ##end legacy drawing
      end

      row_data = files[j].css('sheetData row')

      if(row_data.to_s != "")
        row_data.each do |row|

          unless @data_only
            ##row styles##
            unless row.css('c').nil?
              row_style = '0'
              unless row.attribute('s').nil?
                row_style = row.attribute('s').value.to_s
              end

              wb.worksheets[i].row_styles[row.attribute('r').to_s] = { :style => row_style.to_s  }

              unless row.attribute('ht').to_s == ""
                wb.worksheets[i].change_row_height(Integer(row.attribute('r').to_s)-1,
                  Float(row.attribute('ht').to_s))
              end
            end
            ##end row styles##
          end

          row = row.css('c')
          row.each do |value|
            cell_index = Parser.convert_to_index(value.attribute('r').to_s)
            style_index = nil

            data_type = value.attribute('t').to_s
            
            if (value.css('v').to_s == "") || (value.css('v').children.to_s == "") #no data
              cell_data = nil
            elsif data_type == 's' #shared string
              str_index = Integer(value.css('v').children.to_s)
              cell_data = shared_strings[str_index].to_s
            elsif data_type=='str' #raw string
              cell_data = value.css('v').children.to_s
            elsif data_type=='e' #error
              cell_data = value.css('v').children.to_s
            else# (value.css('v').to_s != "") && (value.css('v').children.to_s != "") #is number
              data_type = ''
              if(value.css('v').children.to_s =~ /\./) #is float
                cell_data = Float(value.css('v').children.to_s)
              else           
                cell_data = Integer(value.css('v').children.to_s)
              end
            end
            cell_formula = nil
            fmla_css = value.css('f')
            if(fmla_css.to_s != "")
              cell_formula = fmla_css.children.to_s
              cell_formula_attr = {}
              cell_formula_attr['t'] = fmla_css.attribute('t').to_s if fmla_css.attribute('t')
              cell_formula_attr['ref'] = fmla_css.attribute('ref').to_s if fmla_css.attribute('ref')
              cell_formula_attr['si'] = fmla_css.attribute('si').to_s if fmla_css.attribute('si')
            end

            unless @data_only
              style_index = value['s'].to_i #nil goes to 0 (default)
            else
              style_index = 0
            end

            wb.worksheets[i].sheet_data[cell_index[0]][cell_index[1]] =
              Cell.new(wb.worksheets[i],cell_index[0],cell_index[1],cell_data,cell_formula,
                data_type,style_index,cell_formula_attr)
            cell = wb.worksheets[i].sheet_data[cell_index[0]][cell_index[1]]
          end
        end
      end
    end

    def Parser.decompress(file_path)
      #ensures it is an xlsx/xlsm file
      if(file_path =~ /(.+)\.xls(x|m)/)
        dir_path = $1.to_s
      else
        raise 'Not .xlsx or .xlsm excel file'
      end

      dir_path = File.join(File.dirname(dir_path), make_safe_name(Time.now.to_s))
      #copies excel file to zip file in same directory
      zip_path = dir_path + '.zip'

      FileUtils.cp(file_path,zip_path)

      MyZip.new.unzip(zip_path,dir_path)
      File.delete(zip_path)

      files = Hash.new

      files['app'] = Nokogiri::XML.parse(File.read(File.join(dir_path,'docProps','app.xml')))
      files['core'] = Nokogiri::XML.parse(File.read(File.join(dir_path,'docProps','core.xml')))

      files['workbook'] = Nokogiri::XML.parse(File.read(File.join(dir_path,'xl','workbook.xml')))

      if(File.exist?(File.join(dir_path,'xl','sharedStrings.xml')))
        files['sharedString'] = Nokogiri::XML.parse(File.read(File.join(dir_path,'xl','sharedStrings.xml')))
      end

      unless @data_only
        #preserves external links
        if File.directory?(File.join(dir_path,'xl','externalLinks'))
          files['externalLinks'] = {}
          ext_links_path = File.join(dir_path,'xl','externalLinks')
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
          drawings_path = File.join(dir_path,'xl','drawings')

          dir = Dir.new(drawings_path).entries.reject {|f| [".", "..", ".DS_Store"].include? f}
          dir.each_with_index do |draw,i|
            files['drawings'][i+1] = File.read(File.join(drawings_path,draw))
          end
        end

        if File.directory?(File.join(dir_path,'xl','printerSettings'))
          files['printerSettings'] = {}
          printer_path = File.join(dir_path,'xl','printerSettings')

          dir = Dir.new(printer_path).entries.reject {|f| [".","..",".DS_Store"].include? f}

          dir.each_with_index do |print, i|
            files['printerSettings'][i+1] = File.open(File.join(printer_path,print), 'rb').read
          end
        end

        if File.directory?(File.join(dir_path,"xl",'worksheets','_rels'))
          files['worksheetRels'] = {}
          worksheet_rels_path = File.join(dir_path,'xl','worksheets','_rels')

          dir = Dir.new(worksheet_rels_path).entries.reject {|f| [".","..",".DS_Store"].include? f}
          dir.each_with_index do |rel, i|
            files['worksheetRels'][i+1] = File.read(File.join(worksheet_rels_path,rel))
          end
        end

        if File.exist?(File.join(dir_path,'xl','vbaProject.bin'))
          files['vbaProject'] = File.open(File.join(dir_path,"xl","vbaProject.bin"),'rb').read
        end

        files['styles'] = Nokogiri::XML.parse(File.read(File.join(dir_path,'xl','styles.xml')))
      end

      @num_sheets = files['workbook'].css('sheets').children.size
      @num_sheets = Integer(@num_sheets)

      #adds all worksheet xml files to files hash
      i=1
      1.upto(@num_sheets) do
        filename = 'sheet'+i.to_s
        files[i] = Nokogiri::XML.parse(File.read(File.join(dir_path,'xl','worksheets',filename+'.xml')))
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
