require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class WorksheetWriter < GenericWriter

    def initialize(workbook, sheet_index = 0)
      @workbook = workbook
      @sheet_index = sheet_index
      @worksheet = @workbook.worksheets[@sheet_index]
    end

    def filepath
      File.join('xl', 'worksheets', "sheet#{@sheet_index + 1}.xml")
    end

    def write()
      render_xml do |xml|
        xml << (xml.create_element('worksheet', 
                  'xmlns'    => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
                  'xmlns:r'  => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
                  'xmlns:mc' => 'http://schemas.openxmlformats.org/markup-compatibility/2006',
                  'xmlns:mv' => 'urn:schemas-microsoft-com:mac:vml',
                  'mc:Ignorable' => 'mv',
                  'mc:PreserveAttributes' => 'mv:*') { |root|

          col = @worksheet.sheet_data.max_by{ |row| row.size }.size
          row = @worksheet.sheet_data.size

          root << xml.create_element('dimension', { :ref => "A1:#{Cell.ind2ref(row - 1, col - 1)}" })

          root << (xml.create_element('sheetViews') { |sheet_views|
            view = @worksheet.sheet_view || {}
            view = view[:attributes] || {}

            sheet_views << (xml.create_element('sheetView', { 
                              :tabSelected     => view[:tabSelected]     || 1,
                              :view            => view[:view]            || 'normalLayout',
                              :workbookViewId  => view[:workbookViewId]  || 0,
                              :zoomScale       => view[:zoomScale]       || 100,
                              :zoomScaleNormal => view[:zoomScaleNormal] || 100 }) {

            #TODO
            #can't be done unless I figure out a way to programmatically add attributes.
            #(can't put xSplit with an invalid value)
            # unless @worksheet.pane.nil?
              # xml.pane('state'=>@worksheet.pane[:state])
            # end
              # unless view[:selection].nil?
                # xml.
              # end

            })
          })

          root << xml.create_element('sheetFormatPr', { :baseColWidth => 10, :defaultRowHeight => 13 })

          ranges = @worksheet.column_ranges
          unless ranges.nil? || ranges.empty?
            root << (xml.create_element('cols') { |cols|
              ranges.each do |range|

                col_attrs = { :min   => range.min + 1,
                              :max   => range.max + 1,
                              :width => range.width || 10,
                              :customWidth => range.custom_width || 0 }

                style_index = @workbook.style_corrector[range.style_index.to_s]
                col_attrs[:style] = style_index if style_index
                cols << (xml.create_element('col', col_attrs))
              end
            })
          end

          root << (xml.create_element('sheetData') { |data|
            @worksheet.sheet_data.each_with_index { |row, i|
              #TODO fix this spans thing. could be 2:3 (not necessary)
              if @worksheet.row_styles[(i+1).to_s].nil?
                @worksheet.row_styles[(i+1).to_s] = {}
                @worksheet.row_styles[(i+1).to_s][:style] = '0'
              end
              custom_format = '1'

              if @worksheet.row_styles[(i+1).to_s][:style].to_s == '0'
                custom_format = '0'
              end

              @worksheet.row_styles[(i+1).to_s][:style] = @workbook.style_corrector[@worksheet.row_styles[(i+1).to_s][:style].to_s]
              row_opts = {
                :r            => i + 1,
                :spans        => "1:#{row.size}",
                :customFormat => custom_format
              }

              unless @worksheet.row_styles[(i+1).to_s][:style].to_s == ''
                row_opts[:s] = @worksheet.row_styles[(i+1).to_s][:style]
              end

              unless @worksheet.row_styles[(i+1).to_s][:height].to_s == ''
                row_opts[:ht] = @worksheet.row_styles[(i+1).to_s][:height]
              end

              unless @worksheet.row_styles[(i+1).to_s][:customheight].to_s == ''
                row_opts[:customHeight] = @worksheet.row_styles[(i+1).to_s][:customHeight]
              end

              data << (xml.create_element('row', row_opts) { |row_xml|
                row.each_with_index { |cell, j|
                  unless cell.nil?
                    #TODO do xml.c for all cases, inside specific.
                    # if cell.formula.nil?
                    cell.style_index = @workbook.style_corrector[cell.style_index.to_s]
                    c_opts = { :r => Cell.ind2ref(i, j), :s => cell.style_index.to_s }

                    unless cell.datatype.nil? || cell.datatype == ''
                      c_opts[:t] = cell.datatype
                    end

                    row_xml << (xml.create_element('c', c_opts) { |cell_xml|
                      unless cell.formula.nil?
                        cell_xml << xml.create_element('f', { 
                            :t   => cell.formula_attributes['t'],
                            :ref => cell.formula_attributes['ref'],
                            :si  => cell.formula_attributes['si'] }, cell.formula )
                      end

                      cell_value = if (cell.datatype == RubyXL::Cell::SHARED_STRING) then
                                     @workbook.shared_strings.get_index(cell.value).to_s
                                   else cell.value
                                   end

                      cell_xml << xml.create_element('v', cell_value)
                    })
                  end #unless cell.nil?
                } #row.each_with_index
              })
            }
          })

          root << xml.create_element('sheetCalcPr', { :fullCalcOnLoad => 1 })

          merged_cells = @worksheet.merged_cells
          unless merged_cells.empty?
            root << xml.create_element('mergeCells', { :count => merged_cells.size }) { |mc|
              @worksheet.merged_cells.each { |ref| mc << xml.create_element('mergeCell', { 'ref' => ref }) }
            }
          end

          root << xml.create_element('phoneticPr', { :fontId => 1, :type => 'noConversion' })

          unless @worksheet.validations.nil?
            root << (xml.create_element('dataValidations', { :count =>@worksheet.validations.size }) { |vals|

              @worksheet.validations.each { |validation|
                attrs = validation[:attributes]

                vals << (xml.create_element('dataValidation', {
                           :type       => attrs[:type],
                           :sqref      => attrs[:sqref],
                           :allowBlank => Integer(attrs[:allowBlank]),
                           :showInputMessage => Integer(attrs[:showInputMessage]),
                           :showErrorMessage => Integer(attrs[:showErrorMessage]) }) { |val|
 
                  unless validation[:formula1].nil?
                    val << val.create_element(:formula1, validation[:formula1])
                  end
                })
              }
            })
          end

          root << xml.create_element('pageMargins', { :left => 0.75, :right => 0.75, :top => 1, :bottom => 1, 
                                                      :header => 0.5, :footer => 0.5 })
          root << xml.create_element('pageSetup', { :orientation => 'portrait',
                                                    :horizontalDpi => 4294967292, :verticalDpi => 4294967292 })

          unless @worksheet.legacy_drawing.nil?
            root << xml.create_element(:legacyDrawing, { 'r:id' => @worksheet.legacy_drawing[:attributes][:id] })
          end

          unless @worksheet.extLst.nil?
            root << (xml.create_element('extLst') { |extlst|
              extlst << (xml.create_element('ext', {
                          'xmlns:mx' => 'http://schemas.microsoft.com/office/mac/excel/2008/main',
                          'uri'      => 'http://schemas.microsoft.com/office/mac/excel/2008/main' }) { |ext|
                ext << xml.create_element('mx:PLV', { :Mode => 1, :OnePage => 0, :WScale => 0 })
              })
            })
          end

          @worksheet.drawings.each { |d| root << xml.create_element('drawing', { 'r:id' => d }) }

        })
      end

    end

  end # class

end
end