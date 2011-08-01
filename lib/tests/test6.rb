require '../rubyXL'
require 'rubygems'
# require File.expand_path(File.join(File.dirname(__FILE__),'workbook'))
# require File.expand_path(File.join(File.dirname(__FILE__),'parser'))
# require File.expand_path(File.join(File.dirname(__FILE__),'color'))
# require File.expand_path(File.join(File.dirname(__FILE__),'cell'))

module RubyXL
wb = Workbook.new([],nil)
wb.worksheets = [Worksheet.new(wb,'Sheet1')]
cell = wb.worksheets[0].add_cell(0,0,'1.00.0')
# cell.change_font_italics(false)
# p cell.is_italicized(wb)
# cell.change_font_bold(false)
# p cell.is_bolded(wb)
# cell.change_font_underline(false)
# p cell.is_underlined(wb)
# p cell.font_name(wb)
# p cell.font_size(wb)
# p cell.font_color(wb)
# p cell.fill_color(wb)

wb.worksheets[0].add_cell(2,5,'$1,000.00')
wb.worksheets[0].add_cell(3,3,'6/14/11')
wb.worksheets[0].add_cell(4,0,1)
wb.worksheets[0].add_cell(4,1,2)
wb.worksheets[0].add_cell(4,2,3)
wb.worksheets[0].add_cell(4,3,4)
wb.worksheets[0].add_cell(4,4,0,'AVERAGE(A5:D5)')

cell.change_font_color('ff0000')

wb.write('/Users/vbhagwat/Desktop/test2/Output/nums.xlsx')


wb2 = Parser.parse('/Users/vbhagwat/Desktop/5-1_5-20.xlsx')
# wb2.worksheets[0].merge_cells(0,1,0,2)
# wb2.worksheets[0].merge_cells(0,0,0,1)
#.change_font_size(30)
wb2.write('/Users/vbhagwat/Desktop/test2/Output/nums2.xlsx')

wb3 = Parser.parse('/Users/vbhagwat/Documents/excelTestFiles/paneWorkbook2/paneWorkbook2.xlsx')


 
# p Color.find(8)
# c = Color::ColorProperties
# p c.has_value?({:hex=>'#000000', :name=>"black"})
# p Color.find('black')
end