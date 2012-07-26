require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class WorksheetWriter
    attr_accessor :dirpath, :filepath, :sheet_num, :workbook, :worksheet

    def initialize(dirpath, wb, sheet_num=1)
      @dirpath = dirpath
      @workbook = wb
      @sheet_num = sheet_num
      @worksheet = @workbook.worksheets[@sheet_num]
      @filepath = dirpath + '/xl/worksheets/sheet'+(sheet_num+1).to_s()+'.xml'
    end

    def write()
      builder = Nokogiri::XML::Builder.new do |xml|
        xml.worksheet('xmlns'=>"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
        'xmlns:r'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        'xmlns:mc'=>"http://schemas.openxmlformats.org/markup-compatibility/2006",
        'xmlns:mv'=>"urn:schemas-microsoft-com:mac:vml",
        'mc:Ignorable'=>'mv',
        'mc:PreserveAttributes'=>'mv:*') {
          col = 0
          @worksheet.sheet_data.each do |row|
            if row.size > col
              col = row.size
            end
          end
          row = Integer(@worksheet.sheet_data.size)
          dimension = 'A1:'
          dimension += Cell.convert_to_cell(row-1,col-1)
          xml.dimension('ref'=>dimension)
          xml.sheetViews {
            view = @worksheet.sheet_view
            if view.nil? || view[:attributes].nil?
              xml.sheetView('tabSelected'=>1,'view'=>'normalLayout','workbookViewId'=>0)
              raise 'end test' #TODO remove
            else
              view = view[:attributes]
              if view[:view].nil?
                view[:view] = 'normalLayout'
              end

              if view[:workbookViewId].nil?
                view[:workbookViewId] = 0
              end

              if view[:zoomScale].nil? || view[:zoomScaleNormal].nil?
                view[:zoomScale] = 100
                view[:zoomScaleNormal] = 100
              end

              if view[:tabSelected].nil?
                view[:tabSelected] = 1
              end

              xml.sheetView('tabSelected'=>view[:tabSelected],
                            'view'=>view[:view],
                            'workbookViewId'=>view[:workbookViewId],
                            'zoomScale'=>view[:zoomScale],
                            'zoomScaleNormal'=>view[:zoomScaleNormal]) {
                #TODO
                #can't be done unless I figure out a way to programmatically add attributes.
                #(can't put xSplit with an invalid value)
                # unless @worksheet.pane.nil?
                  # xml.pane('state'=>@worksheet.pane[:state])
                # end
                  # unless view[:selection].nil?
                    # xml.
                  # end
              }
            end
          }
          xml.sheetFormatPr('baseColWidth'=>'10','defaultRowHeight'=>'13')

          unless @worksheet.cols.nil? || @worksheet.cols.size==0
            xml.cols {
              if !@worksheet.cols.is_a?(Array)
                @worksheet.cols = [@worksheet.cols]
              end
              @worksheet.cols.each do |col|
                if col[:attributes][:customWidth].nil?
                  col[:attributes][:customWidth] = '0'
                end
                if col[:attributes][:width].nil?
                  col[:attributes][:width] = '10'
                  col[:attributes][:customWidth] = '0'
                end
                # unless col[:attributes] == {}
                  xml.col('style'=>@workbook.style_corrector[col[:attributes][:style].to_s].to_s,
                    'min'=>col[:attributes][:min].to_s,
                    'max'=>col[:attributes][:max].to_s,
                    'width'=>col[:attributes][:width].to_s,
                    'customWidth'=>col[:attributes][:customWidth].to_s)
                # end
              end
            }
          end

          xml.sheetData {
            i=0
            @worksheet.sheet_data.each_with_index do |row,i|
              #TODO fix this spans thing. could be 2:3 (not necessary)
              if @worksheet.row_styles[(i+1).to_s].nil?
                @worksheet.row_styles[(i+1).to_s] = {}
                @worksheet.row_styles[(i+1).to_s][:style] = '0'
              end
              custom_format = '1'
              if @worksheet.row_styles[(i+1).to_s][:style] == '0'
                custom_format = '0'
              end

              @worksheet.row_styles[(i+1).to_s][:style] = @workbook.style_corrector[@worksheet.row_styles[(i+1).to_s][:style].to_s]
              row_opts = {
                'r'=>(i+1).to_s,
                'spans'=>'1:'+row.size.to_s,
                'customFormat'=>custom_format
              }
              unless @worksheet.row_styles[(i+1).to_s][:style].to_s == ''
                row_opts['s'] = @worksheet.row_styles[(i+1).to_s][:style].to_s
              end
              unless @worksheet.row_styles[(i+1).to_s][:height].to_s == ''
                row_opts['ht'] = @worksheet.row_styles[(i+1).to_s][:height].to_s
              end
              unless @worksheet.row_styles[(i+1).to_s][:customheight].to_s == ''
                row_opts['customHeight'] = @worksheet.row_styles[(i+1).to_s][:customHeight].to_s
              end
              xml.row(row_opts) {
                row.each_with_index do |dat, j|
                  unless dat.nil?
                      #TODO do xml.c for all cases, inside specific.
                      # if dat.formula.nil?
                      dat.style_index = @workbook.style_corrector[dat.style_index.to_s]
                      c_opts = {'r'=>Cell.convert_to_cell(i,j), 's'=>dat.style_index.to_s}
                      unless dat.datatype.nil? || dat.datatype == ''
                        c_opts['t'] = dat.datatype
                      end
                      xml.c(c_opts) {
                        unless dat.formula.nil?
                          if dat.formula_attributes.nil? || dat.formula_attributes.empty?
                            xml.f dat.formula.to_s
                          else
                            xml.f('t'=>dat.formula_attributes['t'].to_s, 'ref'=>dat.formula_attributes['ref'], 'si'=>dat.formula_attributes['si']).nokogiri dat.formula
                          end
                        end
                        if(dat.datatype == 's')
                          unless dat.value.nil? #empty cell, but has a style
                            xml.v @workbook.shared_strings[dat.value].to_s
                          end
                        elsif(dat.datatype == 'str')
                          xml.v dat.value.to_s
                        elsif(dat.datatype == '') #number
                          xml.v dat.value.to_s
                        end
                      }
                      #
                      # else
                      #   xml.c('r'=>Cell.convert_to_cell(i,j)) {
                      #     xml.v dat.value.to_s
                      #   }
                      # end #data.formula.nil?
                  end #unless dat.nil?
                end #row.each_with_index
              }
            end
          }


          xml.sheetCalcPr('fullCalcOnLoad'=>'1')

          unless @worksheet.merged_cells.nil? || @worksheet.merged_cells.size==0
            #There is some kind of bug here that when merged_cells is sometimes a hash and not an array
            #initial attempt at a fix fails in corrupted excel documents leaving as is for now.
            xml.mergeCells {
              @worksheet.merged_cells.each do |merged_cell|
                xml.mergeCell('ref'=>merged_cell[:attributes][:ref].to_s)
              end
            }
          end

          xml.phoneticPr('fontId'=>'1','type'=>'noConversion')

          unless @worksheet.validations.nil?
            xml.dataValidations('count'=>@worksheet.validations.size.to_s) {
              @worksheet.validations.each do |validation|
                xml.dataValidation('type'=>validation[:attributes][:type],
                  'sqref'=>validation[:attributes][:sqref],
                  'allowBlank'=>Integer(validation[:attributes][:allowBlank]).to_s,
                  'showInputMessage'=>Integer(validation[:attributes][:showInputMessage]).to_s,
                  'showErrorMessage'=>Integer(validation[:attributes][:showErrorMessage]).to_s) {
                    unless validation[:formula1].nil?
                      xml.formula1 validation[:formula1]
                    end
                  }
              end
            }
          end

          xml.pageMargins('left'=>'0.75','right'=>'0.75','top'=>'1',
            'bottom'=>'1','header'=>'0.5','footer'=>'0.5')

          xml.pageSetup('orientation'=>'portrait',
            'horizontalDpi'=>'4294967292', 'verticalDpi'=>'4294967292')

          unless @worksheet.legacy_drawing.nil?
            xml.legacyDrawing('r:id'=>@worksheet.legacy_drawing[:attributes][:id])
          end

          unless @worksheet.extLst.nil?
            xml.extLst {
              xml.ext('xmlns:mx'=>"http://schemas.microsoft.com/office/mac/excel/2008/main",
              'uri'=>"http://schemas.microsoft.com/office/mac/excel/2008/main") {
                xml['mx'].PLV('Mode'=>'1', 'OnePage'=>'0','WScale'=>'0')
              }
            }
          end
        }
      end
      contents = builder.to_xml
      contents = contents.gsub(/\n/,'')
      contents = contents.gsub(/>(\s)+</,'><')
      contents = contents.sub(/<\?xml version=\"1.0\"\?>/,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+"\n")
      contents
    end

  end
end
end
