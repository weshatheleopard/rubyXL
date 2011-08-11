require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class StylesWriter
    attr_accessor :dirpath, :filepath, :workbook

    def initialize(dirpath, wb)
      @dirpath = dirpath
      @workbook = wb
      @filepath = @dirpath + '/xl/styles.xml'
    end

    def write()
      font_id_corrector = {}
      fill_id_corrector = {}
      border_id_corrector = {}
      style_id_corrector = {}

      builder = Nokogiri::XML::Builder.new do |xml|
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
          #default font should stay the same
          font_id_corrector['0']=0
          1.upto(@workbook.fonts.size-1) do |i|
            font_id_corrector[i.to_s] = i-offset
            if @workbook.fonts[i.to_s][:count] == 0
              @workbook.fonts[i.to_s] = nil
              font_id_corrector[i.to_s] = nil
              offset += 1
            end
          end


          offset = 0
          #STARTS AT 2 because excel is stupid
          #and it seems to hard code access the first
          #2 styles.............
          fill_id_corrector['0']=0
          fill_id_corrector['1']=1
          2.upto(@workbook.fills.size-1) do |i|
            fill_id_corrector[i.to_s] = i-offset
            if @workbook.fills[i.to_s][:count] == 0
              @workbook.fills[i.to_s] = nil
              fill_id_corrector[i.to_s] = nil
              offset += 1
            end
          end

          #sets index to itself as init correction
          #if items deleted, indexes adjusted
          #that id 'corrects' to nil
          offset = 0

          #default border should stay the same
          border_id_corrector['0'] = 0
          1.upto(@workbook.borders.size-1) do |i|
            border_id_corrector[i.to_s] = i-offset
            if @workbook.borders[i.to_s][:count] == 0
              @workbook.borders[i.to_s] = nil
              border_id_corrector[i.to_s] = nil
              offset += 1
            end
          end

          if !@workbook.cell_xfs[:xf].is_a?(Array)
            @workbook.cell_xfs[:xf] = [@workbook.cell_xfs[:xf]]
          end
          

          
          style_id_corrector['0']=0
          delete_list = []
          i = 1
          while(i < @workbook.cell_xfs[:xf].size) do
            if style_id_corrector[i.to_s].nil?
              style_id_corrector[i.to_s]= i
            end
            # style correction commented out until bug is fixed
            j = i+1
            while(j < @workbook.cell_xfs[:xf].size) do
              if hash_equal(@workbook.cell_xfs[:xf][i],@workbook.cell_xfs[:xf][j]) #check if this is working
                style_id_corrector[j.to_s] = i
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
              if style_id_corrector[i.to_s] == i
                style_id_corrector[i.to_s] -= offset# unless style_id_corrector[i.to_s].nil? #173 should equal 53, not 52?
              end

              i += 1
            end
            @workbook.cell_xfs[:xf].delete_at(delete_index)
            offset += 1
          end
          
          @workbook.style_corrector = style_id_corrector


          xml.fonts('count'=>@workbook.fonts.size) {
            0.upto(@workbook.fonts.size-1) do |i|
              font = @workbook.fonts[i.to_s]
              unless font.nil?
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
            end
          }

          xml.fills('count'=>@workbook.fills.size) {
            0.upto(@workbook.fills.size-1) do |i|
              fill = @workbook.fills[i.to_s]
              unless fill.nil?
                fill = fill[:fill]
                xml.fill {
                  xml.patternFill('patternType'=>fill[:patternFill][:attributes][:patternType].to_s) {
                    unless fill[:patternFill][:fgColor].nil?
                      fgColor = fill[:patternFill][:fgColor][:attributes]
                      unless fgColor[:indexed].nil?
                        xml.fgColor('indexed'=>fgColor[:indexed].to_s)
                      else
                        unless fgColor[:rgb].nil?
                          xml.fgColor('rgb'=>fgColor[:rgb])
                        end
                      end
                    end
                    unless fill[:patternFill][:bgColor].nil?
                      bgColor = fill[:patternFill][:bgColor][:attributes]
                      unless bgColor[:indexed].nil?
                        xml.bgColor('indexed'=>bgColor[:indexed].to_s)
                      else
                        unless bgColor[:rgb].nil?
                          xml.bgColor('rgb'=>bgColor[:rgb])
                        end
                      end
                    end
                  }
                }
              end
            end
          }

          xml.borders('count'=>@workbook.borders.size) {
            0.upto(@workbook.borders.size-1) do |i|
              border = @workbook.borders[i.to_s]
              unless border.nil?
                border = border[:border]
                xml.border {
                  if border[:left][:attributes].nil?
                    xml.left
                  else
                    xml.left('style'=>border[:left][:attributes][:style]) {
                      unless border[:left][:color].nil?
                        color = border[:left][:color][:attributes]
                        unless color[:indexed].nil?
                          xml.color('indexed'=>color[:indexed])
                        else
                          unless color[:rgb].nil?
                            xml.color('rgb'=>color[:rgb])
                          end
                        end
                      end
                    }
                  end
                  if border[:right][:attributes].nil?
                    xml.right
                  else
                    xml.right('style'=>border[:right][:attributes][:style]) {
                      unless border[:right][:color].nil?
                        color = border[:right][:color][:attributes]
                        unless color[:indexed].nil?
                          xml.color('indexed'=>color[:indexed])
                        else
                          unless color[:rgb].nil?
                            xml.color('rgb'=>color[:rgb])
                          end
                        end
                      end
                    }
                  end
                  if border[:top][:attributes].nil?
                    xml.top
                  else
                    xml.top('style'=>border[:top][:attributes][:style]) {
                      unless border[:top][:color].nil?
                        color = border[:top][:color][:attributes]
                        unless color[:indexed].nil?
                          xml.color('indexed'=>color[:indexed])
                        else
                          unless color[:rgb].nil?
                            xml.color('rgb'=>color[:rgb])
                          end
                        end
                      end
                    }
                  end
                  if border[:bottom][:attributes].nil?
                    xml.bottom
                  else
                    xml.bottom('style'=>border[:bottom][:attributes][:style]) {
                      unless border[:bottom][:color].nil?
                        color = border[:bottom][:color][:attributes]
                        unless color[:indexed].nil?
                          xml.color('indexed'=>color[:indexed])
                        else
                          unless color[:rgb].nil?
                            xml.color('rgb'=>color[:rgb])
                          end
                        end
                      end
                    }
                  end
                  if border[:diagonal][:attributes].nil?
                    xml.diagonal
                  else
                    xml.diagonal('style'=>border[:diagonal][:attributes][:style]) {
                      unless border[:diagonal][:color].nil?
                        color = border[:diagonal][:color][:attributes]
                        unless color[:indexed].nil?
                          xml.color('indexed'=>color[:indexed])
                        else
                          unless color[:rgb].nil?
                            xml.color('rgb'=>color[:rgb])
                          end
                        end
                      end
                    }
                  end
                }
              end #unless border.nil?
            end #0.upto(size)
          }

          xml.cellStyleXfs('count'=>@workbook.cell_style_xfs[:attributes][:count].to_s) {
            @workbook.cell_style_xfs[:xf].each do |style|
              style = @workbook.get_style_attributes(style)
              xml.xf('numFmtId'=>style[:numFmtId].to_s,
              'fontId'=>font_id_corrector[style[:fontId].to_s].to_s,
              'fillId'=>fill_id_corrector[style[:fillId].to_s].to_s,
              'borderId'=>border_id_corrector[style[:borderId].to_s].to_s)
            end
          }

          xml.cellXfs('count'=>@workbook.cell_xfs[:xf].size) {
            @workbook.cell_xfs[:xf].each do |xf_obj|
              xf = @workbook.get_style_attributes(xf_obj)

              xml.xf('numFmtId'=>xf[:numFmtId].to_s,
              'fontId'=>font_id_corrector[xf[:fontId].to_s].to_s,
              'fillId'=>fill_id_corrector[xf[:fillId].to_s].to_s,
              'borderId'=>border_id_corrector[xf[:borderId].to_s].to_s,
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
          xml.cellStyles('count'=>@workbook.cell_styles[:attributes][:count].to_s) {

            @workbook.cell_styles[:cellStyle].each do |style|
              style = @workbook.get_style_attributes(style)
              xml.cellStyle('name'=>style[:name].to_s,
              'xfId'=>style[:xfId].to_s,
              'builtinId'=>style[:builtinId].to_s)
            end
          }
          xml.dxfs('count'=>'0')
          xml.tableStyles('count'=>'0', 'defaultTableStyle'=>'TableStyleMedium9')

          unless @colors.nil?
            xml.colors {
              unless @colors[:indexedColors].nil?
                xml.indexedColors {
                  @colors[:indexedColors].each do |rgb_color|
                    xml.rgbColor rgb_color[:attributes][:rgb]
                  end
                }
              end

              unless @colors[:mruColors].nil?
                xml.mruColors {
                  @colors[:mruColors].each do |color|
                    xml.color color[:attributes][:rgb]
                  end
                }
              end
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
