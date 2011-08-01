require File.expand_path(File.join(File.dirname(__FILE__),'workbook'))
require File.expand_path(File.join(File.dirname(__FILE__),'worksheet'))
require File.expand_path(File.join(File.dirname(__FILE__),'cell'))
require File.expand_path(File.join(File.dirname(__FILE__),'color'))

module RubyXL
class HtmlGenerator
  def generate(wb)
    puts 'generating html....'
    html = "<!doctype html><html><head></head><body>"
    html += "<h1 style='text-align:center'>"
    html += wb.filepath
    html += "</h1>"
    sheets = wb.worksheets
    sheets.each do |sheet|
      html += "<h2>"
      html += sheet.sheet_name
      html +="</h2>"
      html += '<table border="1">'
      data = sheet.sheet_data
      data.each do |row|
        html += '<tr>'
        row.each do |val|
          html += '<td'
          # p val
          if(val != nil)
            #font color
            html+= ' style="color:'
            #TODO make this not cast it to int
            html += "\##{val.font_color}"

            #background color
            html+= '; background-color:'
            html += "\##{val.fill_color}"

            if(val.is_italicized == true)
              html+= '; font-style:italic'
            end
            if(val.is_bolded == true)
              html+= '; font-weight:bold'
            end
            if(val.is_underlined == true)
              html+= '; text-decoration:underline'
            end

            html += '"'
          end
          html += '>'
          if(val != nil && val.value != nil)
            if(val.formula != nil)
              html += val.formula + ' = '
            end
            html += val.value.to_s()
          else
            html += "----"
          end
          html += '</td>'
        end
        html += '</tr>'
      end
      html += '</table>'
      html += '<br /><br />'
    end


    html += "</body></html>"
    File.open('test.html','w'){|f| f.write(html)}
    puts 'done.'
  end
end
end
