require '../rubyXL'
require 'rubygems'
require 'rubyXL'
# require File.expand_path(File.join(File.dirname(__FILE__),'workbook'))

module RubyXL
wb = Workbook.new([],nil)
wb.worksheets = [Worksheet.new(wb,'Sheet1')]

cell = wb.worksheets[0].add_cell(0,0,'A1')
cell2 = wb.worksheets[0].add_cell(0,1,'B1')
cell3 = wb.worksheets[0].add_cell(0,5,'F1')

wb.worksheets[0].sheet_data[0].each do |c|
  unless c.nil?
    c.change_font_bold(true)
    c.change_font_underline(true)
  end
end

cell.change_horizontal_alignment('center')
wb.worksheets[0].change_row_horizontal_alignment(0,'justify')
wb.worksheets[0].change_row_vertical_alignment(0,'center')

wb.worksheets[0].change_row_fill(0,'FF0000')
cell2.change_fill('0000FF')

cell4 = wb.worksheets[0].add_cell(0,3,'D1')
puts '



'
p wb.fills
puts '




'
cell4.change_fill('A43502')

cell5 = wb.worksheets[0].add_cell(1,0,'A2')
wb.worksheets[0].change_row_fill(1,'00FF00')
cell5.change_fill('FF0000')

wb.write('/Users/vbhagwat/Desktop/test2/Output/test7.xlsx')
end