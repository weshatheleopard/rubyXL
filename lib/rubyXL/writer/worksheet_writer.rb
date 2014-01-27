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
          root << xml.create_element('sheetFormatPr', { :baseColWidth => 10, :defaultRowHeight => 13 })
          root << xml.create_element('sheetCalcPr', { :fullCalcOnLoad => 1 })
          root << xml.create_element('pageMargins', { :left => 0.75, :right => 0.75, :top => 1, :bottom => 1, 
                                                      :header => 0.5, :footer => 0.5 })
          root << xml.create_element('pageSetup', { :orientation => 'portrait',
                                                    :horizontalDpi => 4294967292, :verticalDpi => 4294967292 })

          unless @worksheet.extLst.nil?
            root << (xml.create_element('extLst') { |extlst|
              extlst << (xml.create_element('ext', {
                          'xmlns:mx' => 'http://schemas.microsoft.com/office/mac/excel/2008/main',
                          'uri'      => 'http://schemas.microsoft.com/office/mac/excel/2008/main' }) { |ext|
                ext << xml.create_element('mx:PLV', { :Mode => 1, :OnePage => 0, :WScale => 0 })
              })
            })
          end

        })
=end

      end
    end

  end # class

end
end