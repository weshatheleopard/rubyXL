module RubyXL
  module Writer
    class StylesWriter < GenericWriter

      def filepath
        File.join('xl', 'styles.xml')
      end

      def write()

        render_xml do |xml|
          xml << @workbook.stylesheet.write_xml(xml)
=begin
          xml << (xml.create_element('styleSheet', :xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main") { |root|
            unless @workbook.num_fmts.empty?
              root << (xml.create_element('numFmts', :count => @workbook.num_fmts.size) { |numfmts|
                @workbook.num_fmts.each { |numfmt| numfmts << numfmt.write_xml(xml) }
              })
            end

            root << (xml.create_element('fonts', :count => @workbook.fonts.size) { |fonts|
              @workbook.fonts.each_with_index { |font, i| fonts << font.write_xml(xml) }
            })

            root << (xml.create_element('fills', :count => @workbook.fills.size) { |fills|
              @workbook.fills.each_with_index { |fill, i| fills << fill.write_xml(xml) }
            })

            root << (xml.create_element('borders', :count => @workbook.borders.size) { |borders|
              @workbook.borders.each_with_index { |border, i| borders << border.write_xml(xml) }
            })

            root << (xml.create_element('cellStyleXfs', :count => @workbook.cell_style_xfs.count) { |cxfs|
              @workbook.cell_style_xfs.each { |xf| cxfs << xf.write_xml(xml) }
            })


            root << (xml.create_element('cellXfs', :count => @workbook.cell_xfs.size) { |cxfs|
              @workbook.cell_xfs.each { |xf| cxfs << xf.write_xml(xml) }
            })

            root << (xml.create_element('cellStyles', :count => @workbook.cell_styles.size) { |cell_styles|
              @workbook.cell_styles.each { |style| cell_styles << style.write_xml(xml) }
            })

            root << xml.create_element('dxfs', :count => 0)
            root << xml.create_element('tableStyles', :count => 0, :defaultTableStyle => 'TableStyleMedium9')

            unless @workbook.colors.empty?
              root << (xml.create_element('colors') { |colors|
                @workbook.colors.each_pair { |color_type, color_array|
                  colors << (xml.create_element(color_type) { |type|
                    color_array.each { |color| type << color.write_xml(xml) }
                  })
                }
              })
            end

          })
=end
        end
      end

    end
  end
end
