require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class StylesWriter < GenericWriter

    def filepath
      File.join('xl', 'styles.xml')
    end

    def write()
      @style_id_corrector = {}

      @font_id_corrector = []  # Relation "old_index -> corrected_index"
      corrected_index = 0

      @workbook.fonts.each_with_index { |font, i|
        if font.count > 0 || i == 0 then # Default font #0 should stay the same
          @font_id_corrector[i] = corrected_index
          corrected_index += 1
        end
      }

      @fill_id_corrector = []
      corrected_index = 0
      @workbook.fills.each_with_index { |fill, i|
        if fill.count > 0 || i < 2 then # First 2 fills are hardcoded
          @fill_id_corrector[i] = corrected_index
          corrected_index += 1
        end
      }

      @border_id_corrector = []
      corrected_index = 0
      @workbook.borders.each_with_index { |border, i|
        if (border.count > 0) || (i == 0) then # Keep special style #0 and all used styles
          @border_id_corrector[i] = corrected_index
          corrected_index += 1
        end
      }

      @workbook.cell_xfs[:xf] = [@workbook.cell_xfs[:xf]] unless @workbook.cell_xfs[:xf].is_a?(Array)

      @style_id_corrector[0] = 0
      delete_list = []
      i = 1
      while(i < @workbook.cell_xfs[:xf].size) do
        if @style_id_corrector[i].nil?
          @style_id_corrector[i]= i
        end
        # style correction commented out until bug is fixed
        j = i+1
        while(j < @workbook.cell_xfs[:xf].size) do
          if hash_equal(@workbook.cell_xfs[:xf][i],@workbook.cell_xfs[:xf][j]) #check if this is working
            @style_id_corrector[j] = i
            delete_list << j
          end
          j += 1
        end
        i += 1
      end
      
      #go through delete list, if before delete_list index 0, offset 0, if before delete_list index 1, offset 1, etc.
      delete_list.sort!

      i = 1
      offset = 0
      offset_corrector = 0
      delete_list << @workbook.cell_xfs[:xf].size
      while offset < delete_list.size do
        delete_index = delete_list[offset] - offset

        while i <= delete_list[offset] do #if <= instead of <, fixes odd border but adds random cells with fill              
          if @style_id_corrector[i] == i
            @style_id_corrector[i] -= offset# unless @style_id_corrector[i.to_s].nil? #173 should equal 53, not 52?
          end

          i += 1
        end
        @workbook.cell_xfs[:xf].delete_at(delete_index)
        offset += 1
      end
      
      @workbook.style_corrector = @style_id_corrector

      render_xml do |xml|
        xml << (xml.create_element('styleSheet', :xmlns => "http://schemas.openxmlformats.org/spreadsheetml/2006/main") { |root|
          unless @workbook.num_fmts.empty?
            root << (xml.create_element('numFmts', :count => @workbook.num_fmts.size) { |numfmts|
              @workbook.num_fmts.each { |numfmt| numfmts << numfmt.write_xml(xml) }
            })
          end

          root << (xml.create_element('fonts', :count => @workbook.fonts.size) { |fonts|
            @workbook.fonts.each_with_index { |font, i| fonts << font.write_xml(xml) unless @font_id_corrector[i].nil? }
          })

          root << (xml.create_element('fills', :count => @workbook.fills.size) { |fills|
            @workbook.fills.each_with_index { |fill, i| fills << fill.write_xml(xml) unless @fill_id_corrector[i].nil? }
          })

          root << (xml.create_element('borders', :count => @workbook.borders.size) { |borders|
            @workbook.borders.each_with_index { |border, i| borders << border.write_xml(xml) unless @border_id_corrector[i].nil? }
          })

          root << (xml.create_element('cellStyleXfs', :count => @workbook.cell_style_xfs[:attributes][:count]) { |cxfs|
            @workbook.cell_style_xfs[:xf].each { |style|
              style = @workbook.get_style_attributes(style)
              cxfs << xml.create_element('xf', :numFmtId  => style[:numFmtId],
                                               :fontId => @font_id_corrector[style[:fontId]],
                                               :fillId => @fill_id_corrector[style[:fillId]],
                                               :borderId => @border_id_corrector[style[:borderId]])
            }
          })


          root << (xml.create_element('cellXfs', :count => @workbook.cell_xfs[:xf].size) { |cxfs|
            @workbook.cell_xfs[:xf].each { |xf_obj|
              xf = @workbook.get_style_attributes(xf_obj)
              cxfs << (xml.create_element('xf', :numFmtId => xf[:numFmtId],
                                                :fontId => @font_id_corrector[xf[:fontId]],
                                                :fillId => @fill_id_corrector[xf[:fillId]],
                                                :borderId => @border_id_corrector[xf[:borderId]],
                                                :xfId => xf[:xfId].to_s,
                                                :applyFont => xf[:applyFont].to_i, #0 if nil
                                                :applyFill => xf[:applyFill].to_i,
                                                :applyAlignment => xf[:applyAlignment].to_i,
                                                :applyNumberFormat => xf[:applyNumberFormat].to_i) { |xf_xml|

                unless xf_obj.is_a?(Array)
                  unless xf_obj[:alignment].nil?
                    xf_xml << xml.create_element('alignment', :horizontal => xf_obj[:alignment][:attributes][:horizontal],
                                                              :vertical => xf_obj[:alignment][:attributes][:vertical],
                                                              :wrapText => xf_obj[:alignment][:attributes][:wrapText])
                  end
                end

              })


            }
          })

          root << (xml.create_element('cellStyles', :count => @workbook.cell_styles.size) { |cell_styles|
            @workbook.cell_styles.each { |style|
              cell_styles << style.write_xml(xml)
            }
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
      end

    end

    private

    def hash_equal(h1,h2)
      if h1.nil?
        if h2.nil?
          return true
        else
          return false
        end
      elsif h2.nil?
        return false
      end
      if h1.size != h2.size
        return false
      end

      h1.each do |k,v|
        if (h1[k].is_a?String) || (h2[k].is_a?String)
          if (h1.is_a?Hash) && (h2.is_a?Hash)
            unless hash_equal(h1[k].to_s,h2[k].to_s)
              return false
            end
          else
            unless h1[k].to_s == h2[k].to_s
              return false
            end
          end
        else
          unless h1[k] == h2[k]
            return false
          end
        end
      end

      true
    end

  end
end
end
