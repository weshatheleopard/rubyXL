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

          min_row = min_col = nil
          max_row = max_col = 0

          @worksheet.sheet_data.rows.each_with_index { |row, row_index|
            next if row.nil?
            row_has_cells = false # Row may have had all its cells deleted.

            row.cells.each_with_index { |cell, col_index|
              next if cell.nil?
              row_has_cells = true
              min_col = col_index if min_col.nil?
              max_col = col_index if col_index > max_col
            }

            # TODO: check for other attributes, as row might have no cells, but row style.
            if row_has_cells then 
              min_row = row_index if min_row.nil?
              max_row = row_index if row_index > min_row
            end
          }

          root << xml.create_element('dimension', { :ref => RubyXL::Reference.new(min_row, max_row, min_col, max_col) })

          unless @worksheet.sheet_views.empty?
            root << xml.create_element('sheetViews') { |sheet_views|
              @worksheet.sheet_views.each { |sheet_view| sheet_views << sheet_view.write_xml(xml) }
            }
          end

          root << xml.create_element('sheetFormatPr', { :baseColWidth => 10, :defaultRowHeight => 13 })

          ranges = @worksheet.column_ranges
          unless ranges.nil? || ranges.empty?
            root << (xml.create_element('cols') { |cols|
              ranges.each { |range| cols << range.write_xml(xml) }
            })
          end

          root << (xml.create_element('sheetData') { |data|
            @worksheet.sheet_data.rows.each_with_index { |row, i|
              next if row.nil?

              row_opts = {
                :r            => i + 1,
                :spans        => "#{min_col + 1}:#{max_col + 1}",
                :customFormat => row.custom_format || (row.s.to_i != 0)
              }

              row_opts[:s] = row.s if row.s
              row_opts[:ht] = row.ht if row.ht
              row_opts[:customHeight] = row.custom_height if row.custom_height

              data << (xml.create_element('row', row_opts) { |row_xml|
                row.cells.each_with_index { |cell, row_index|
                  next if cell.nil?
                  cell.r ||= RubyXL::Reference.new(i, row_index)
                  row_xml << cell.write_xml(xml)
                }
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

          unless @worksheet.validations.empty?
            root << (xml.create_element('dataValidations', { :count => @worksheet.validations.size }) { |validations|
              @worksheet.validations.each { |validation| validations << validation.write_xml(xml) }
            })
          end

          root << xml.create_element('pageMargins', { :left => 0.75, :right => 0.75, :top => 1, :bottom => 1, 
                                                      :header => 0.5, :footer => 0.5 })
          root << xml.create_element('pageSetup', { :orientation => 'portrait',
                                                    :horizontalDpi => 4294967292, :verticalDpi => 4294967292 })

          @worksheet.legacy_drawings.each { |drawing| root << drawing.write_xml(xml) }

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