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
        xml << @worksheet.write_xml(xml)

=begin
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
          root << xml.create_element('sheetFormatPr', { :baseColWidth => 10, :defaultRowHeight => 13 })
          root << xml.create_element('sheetCalcPr', { :fullCalcOnLoad => 1 })
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
=end

      end
    end

  end # class

end
end