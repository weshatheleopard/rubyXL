require '../rubyXL'
require 'rubygems'
require 'nokogiri'
# require File.expand_path(File.join(File.dirname(__FILE__),'Hash'))
# require File.expand_path(File.join(File.dirname(__FILE__),'cell'))
# require File.expand_path(File.join(File.dirname(__FILE__),'workbook'))
# require File.expand_path(File.join(File.dirname(__FILE__),'parser'))

# module RubyXL
file_path = '/Users/vbhagwat/Documents/excelTestFiles/styledWorkbook2/styledWorkbook2.xlsx'
wb = RubyXL::Parser.parse(file_path)
puts "parsed #{file_path}"
p wb.cell_xfs
cell = wb.worksheets[0].sheet_data[0][0]
cell2 = wb.worksheets[0].sheet_data[1][1]
cell3 = wb.worksheets[0].sheet_data[0][2]

cell4 = wb.worksheets[0].add_cell(4,4,'test')
cell5 = wb.worksheets[0].add_cell(4,4,'test2',nil,false)

# cells = Array.new()
# 
# cells << Cell.new(0,3,'black', nil,'str') #8
# cells << Cell.new(0,4,'white',nil,'str')
# cells << Cell.new(0,5,'red',nil,'str')
# cells << Cell.new(0,6,'brightgreen',nil,'str')
# cells << Cell.new(0,7,'blue',nil,'str')
# cells << Cell.new(0,8,'yellow',nil,'str')
# cells << Cell.new(0,9,'magenta',nil,'str')
# cells << Cell.new(0,10,'cyan',nil,'str')
# cells << Cell.new(0,11,'darkred',nil,'str')
# cells << Cell.new(0,12,'green',nil,'str')
# cells << Cell.new(0,13,'darkblue',nil,'str')
# cells << Cell.new(0,14,'darkyellow',nil,'str')
# cells << Cell.new(0,15,'purple',nil,'str')
# cells << Cell.new(0,16,'teal',nil,'str')
# cells << Cell.new(0,17,'gray25',nil,'str')
# cells << Cell.new(0,18,'gray50',nil,'str') #23
# 19.upto(58) do |ind|
#   cells << Cell.new(0,ind,(ind+5).to_s,nil,'str')
# end
# 
# cells.each_with_index do |c,i|
#   wb.worksheets[0].sheet_data[0] << c
#   c.change_fill((i+8).to_s)
# end
def print_stuff(wb)
  puts 'begin print_stuff'
  puts ''
  puts ''
  puts ''
  p wb.cell_xfs[:xf][0]
  puts ''
  puts ''
  puts ''
  puts 'end print_stuff'
end

print_stuff(wb)
cell.change_font_name('Verdana')
print_stuff(wb)
cell2.change_font_size(30)
print_stuff(wb)
cell.change_fill('ff0000')
print_stuff(wb)
wb.worksheets[0].change_row_font_name(1,'Courier')
print_stuff(wb)
wb.worksheets[0].change_row_fill(1,'00ff00')
print_stuff(wb)
wb.worksheets[0].change_column_font_size(0,20)
print_stuff(wb)
cell.change_font_bold(false)
print_stuff(wb)
cell2.change_font_underline(true)
print_stuff(wb)
cell3.change_fill('000000')
print_stuff(wb)
cell3.change_font_name('Verddddana')
print_stuff(wb)
# cell3.change_fill('52')

wb.write('/Users/vbhagwat/Desktop/test2/Output/cell.xlsx')

p wb.style_corrector

print_stuff(wb)
puts 'completed writing /Users/vbhagwat/Desktop/test2/Output/cell.xlsx'
#TODO 

# end