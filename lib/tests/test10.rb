require 'rubygems'
require 'rubyXL'

@workbook  = RubyXL::Workbook.new    
@worksheet = RubyXL::Worksheet.new(@workbook)    
@workbook.worksheets[0] = @worksheet     
(0..10).each do |i|
  (0..10).each do |j|        
    @worksheet.add_cell(i, j, "#{i}:#{j}")
  end
end

@worksheet.change_column_font_color(0,'ff0000')
@worksheet.change_column_font_size(0,"30")

@workbook.write('/Users/vbhagwat/Desktop/test2/Output/test10.xlsx')