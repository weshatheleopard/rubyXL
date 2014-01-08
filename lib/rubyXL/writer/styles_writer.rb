require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class StylesWriter < GenericWriter

    def filepath
      File.join('xl', 'styles.xml')
    end

    def write()
      @fill_id_corrector = {}
      @border_id_corrector = {}
      @style_id_corrector = {}

      @font_id_corrector = []  # Relation "old_index -> corrected_index"
      corrected_index = 0
      @workbook.fonts.each_with_index { |font, i|
        if @workbook.fonts[i][:count] > 0 || i == 0 then # Default font #0 should stay the same
          @font_id_corrector[i] = corrected_index
          corrected_index += 1
        end
      }

      build_xml do |xml|
        xml.styleSheet('xmlns'=>"http://schemas.openxmlformats.org/spreadsheetml/2006/main") {
          unless @workbook.num_fmts.nil? || @workbook.num_fmts[:attributes].nil?
            xml.numFmts('count'=>@workbook.num_fmts[:attributes][:count].to_s) {
              @workbook.num_fmts[:numFmt].each do |fmt|
                attributes = fmt[:attributes]
                xml.numFmt('numFmtId'=>attributes[:numFmtId].to_s,
                  'formatCode'=>attributes[:formatCode].to_s)
              end
            }
          end

          offset = 0
          #STARTS AT 2 because excel is stupid
          #and it seems to hard code access the first
          #2 styles.............
          @fill_id_corrector[0] = 0
          @fill_id_corrector[1] = 1
          2.upto(@workbook.fills.size - 1) { |i|
            @fill_id_corrector[i] = i-offset
            if @workbook.fills[i].count == 0
              @workbook.fills[i] = nil
              @fill_id_corrector[i] = nil
              offset += 1
            end
          }

          offset = 0
          @workbook.borders.each_with_index { |border, i|
            if (i == 0) || (border.count > 0) then # Keep special style #0 and all used styles
              @border_id_corrector[i] = i - offset
            else # Remove all others
              @workbook.borders[i] = @border_id_corrector[i] = nil
              offset += 1
            end
          }

          if !@workbook.cell_xfs[:xf].is_a?(Array)
            @workbook.cell_xfs[:xf] = [@workbook.cell_xfs[:xf]]
          end

          @style_id_corrector['0']=0
          delete_list = []
          i = 1
          while(i < @workbook.cell_xfs[:xf].size) do
            if @style_id_corrector[i.to_s].nil?
              @style_id_corrector[i.to_s]= i
            end
            # style correction commented out until bug is fixed
            j = i+1
            while(j < @workbook.cell_xfs[:xf].size) do
              if hash_equal(@workbook.cell_xfs[:xf][i],@workbook.cell_xfs[:xf][j]) #check if this is working
                @style_id_corrector[j.to_s] = i
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
              if @style_id_corrector[i.to_s] == i
                @style_id_corrector[i.to_s] -= offset# unless @style_id_corrector[i.to_s].nil? #173 should equal 53, not 52?
              end

              i += 1
            end
            @workbook.cell_xfs[:xf].delete_at(delete_index)
            offset += 1
          end
          
          @workbook.style_corrector = @style_id_corrector


          xml.fonts('count'=>@workbook.fonts.size) {
            0.upto(@workbook.fonts.size-1) do |i|
              font = @workbook.fonts[i]
              next if font.nil? || @font_id_corrector[i].nil?
                font = font[:font]

                xml.font {
                  xml.sz('val'=>font[:sz][:attributes][:val].to_s)
                  unless font[:b].nil?
                    xml.b
                  end
                  unless font[:i].nil?
                    xml.i
                  end
                  unless font[:u].nil?
                    xml.u
                  end
                  unless font[:strike].nil?
                    xml.strike
                  end
                  unless font[:color].nil?
                    unless font[:color][:attributes][:indexed].nil?
                      xml.color('indexed'=>font[:color][:attributes][:indexed])
                    else
                      unless font[:color][:attributes][:rgb].nil?
                        xml.color('rgb'=>font[:color][:attributes][:rgb])
                      else
                        unless font[:color][:attributes][:theme].nil?
                          xml.color('theme'=>font[:color][:attributes][:theme])
                        end
                      end
                    end
                  end
                  unless font[:family].nil?
                    xml.family('val'=>font[:family][:attributes][:val].to_s)
                  end
                  unless font[:scheme].nil?
                    xml.scheme('val'=>font[:scheme][:attributes][:val].to_s)
                  end

                  xml.name('val'=>font[:name][:attributes][:val].to_s)
              }
            end
          }

          xml.fills('count' => @workbook.fills.size) {
            @workbook.fills.each { |fill| fill.build_xml(xml) unless fill.nil? }
          }

          xml.borders('count' => @workbook.borders.size) {
            @workbook.borders.each { |border| border.build_xml(xml) unless border.nil? }
          }

          xml.cellStyleXfs('count' => @workbook.cell_style_xfs[:attributes][:count]) {
            @workbook.cell_style_xfs[:xf].each do |style|
              style = @workbook.get_style_attributes(style)
              xml.xf('numFmtId'=>style[:numFmtId].to_s,
              'fontId'=> @font_id_corrector[style[:fontId]],
              'fillId'=>@fill_id_corrector[style[:fillId].to_s].to_s,
              'borderId'=> @border_id_corrector[style[:borderId]])
            end
          }

          xml.cellXfs('count'=>@workbook.cell_xfs[:xf].size) {
            @workbook.cell_xfs[:xf].each do |xf_obj|
              xf = @workbook.get_style_attributes(xf_obj)

              xml.xf('numFmtId'=>xf[:numFmtId].to_s,
              'fontId'=> @font_id_corrector[xf[:fontId]],
              'fillId'=>@fill_id_corrector[xf[:fillId]],
              'borderId' => @border_id_corrector[xf[:borderId]],
              'xfId'=>xf[:xfId].to_s,
              'applyFont'=>xf[:applyFont].to_i.to_s, #0 if nil
              'applyFill'=>xf[:applyFill].to_i.to_s,
              'applyAlignment'=>xf[:applyAlignment].to_i.to_s,
              'applyNumberFormat'=>xf[:applyNumberFormat].to_i.to_s) {
                unless xf_obj.is_a?Array
                  unless xf_obj[:alignment].nil?
                    xml.alignment('horizontal'=>xf_obj[:alignment][:attributes][:horizontal].to_s,
                                  'vertical'=>xf_obj[:alignment][:attributes][:vertical].to_s,
                                  'wrapText'=>xf_obj[:alignment][:attributes][:wrapText].to_s)
                  end
                end
              }
            end
          }
          xml.cellStyles('count'=>@workbook.cell_styles[:attributes][:count]) {

            @workbook.cell_styles[:cellStyle].each do |style|
              style = @workbook.get_style_attributes(style)
              xml.cellStyle('name'=>style[:name].to_s,
              'xfId'=>style[:xfId].to_s,
              'builtinId'=>style[:builtinId].to_s)
            end
          }
          xml.dxfs('count'=>'0')
          xml.tableStyles('count'=>'0', 'defaultTableStyle'=>'TableStyleMedium9')

          unless @workbook.colors.empty?
            xml.colors {
              @workbook.colors.each_pair { |k, v|
                xml.send(k.to_sym) {
                  v.each { |color| color.build_xml(xml) }
                }
              }
            }
          end

        }
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
