require '../rubyXL'
require 'rubygems'
require 'FileUtils'
require 'nokogiri'


module RubyXL
#writes blank_workbook to .xlsx
 
# file_p ath = "/Users/vbhagwat/Documents/excelTestFiles/styledWorkbook/styledWorkbook.xlsx"
# file_path = "/Users/vbhagwat/Documents/excelTestFiles/smallFormulaWorkbook2/smallFormulaWorkbook2.xlsx"
# file_path = "/Users/vbhagwat/Documents/excelTestFiles/threeSheetWorkbook/threeSheetWorkbook.xlsx"
# file_path = "/Users/vbhagwat/Documents/excelTestFiles/numWorkbook2/numWorkbook2.xlsx"
file_path = "/Users/vbhagwat/Documents/excelTestFiles/blankWorkbook/blankWorkbook.xlsx"
# file_path = '/Users/vbhagwat/Desktop/test2/Archive.xlsx'
# file_path = "/Users/vbhagwat/Desktop/5-1_5-20.xlsx"
 
# puts 'begin parsing ' + file_path
# wb = Parser.parse(file_path)
# puts 'completed parsing ' + file_path 

#  def initialize(worksheets,filepath,creator=nil,modifier=nil,created_at=nil,modified_at=nil, company=nil, application=nil,appversion=nil,


dirpath = '/Users/vbhagwat/Documents/excelTestFiles'
wb = Workbook.new(
  [], #worksheets
  file_path, #filepath
  'Vivek Bhagwat', #creator
  'Vivek Bhagwat', #modifier
  '2011-05-16T15:41:00Z', #created_at
  'Gilt Groupe', #company
  'Microsoft Macintosh Excel', #application
  '12.0000')
  wb.worksheets = [Worksheet.new(wb,'Sheet1',[[nil]])]
wb2 = Workbook.new(
  [], #worksheets
  file_path, #filepath
  'Vivek Bhagwat', #creator
  'Vivek Bhagwat', #modifier
  '2011-05-16T15:41:00Z', #created_at
  'Gilt Groupe', #company
  'Microsoft Macintosh Excel', #application
  '12.0000')
  wb.worksheets = [Worksheet.new('Sheet1')]

wb2[0].sheet_data = [[Cell.new(wb2[0],0,0,'6/8/2011'), Cell.new(wb2[0],0,1,'test2')]]

# wb = Parser.parse(file_path)
file_path = '/Users/vbhagwat/Desktop/test2/Output/blank.xlsx'
puts 'begin writing ' + file_path
p wb
wb.write(file_path)
puts 'completed writing ' + file_path

# file_path = "/Users/vbhagwat/Documents/excelTestFiles/twoSheetWorkbook/twoSheetWorkbook.xlsx"
file_path = "/Users/vbhagwat/Documents/excelTestFiles/styledWorkbook7/styledWorkbook7.xlsx"

wb3 = Parser.parse(file_path)
wb3.write('/Users/vbhagwat/Desktop/test2/Output/small_before.xlsx')


# file_path = "/Users/vbhagwat/Desktop/5-1_5-20.xlsx"
wb2 = Parser.parse(file_path)
file_path = '/Users/vbhagwat/Desktop/test2/Output/small.xlsx'

wb2.worksheets[0].change_row_font_name(1,'Courier') #0 indexed.
wb2.worksheets[0].change_row_font_size(1,30) #0 indexed.
cell = wb2.worksheets[0].add_cell(0,0,'A1')
wb2.worksheets[0].change_row_fill(0,'00FF00')
wb2.worksheets[0].add_cell(0,1,'B1')

wb2.worksheets[0].change_column_fill(0,'FFFF00')
wb2.worksheets[0].change_column_font_size(0,20)
wb2.worksheets[0].change_column_width(0,30)
wb2.worksheets[0].change_row_height(2,100)


p wb2.worksheets[0].row_styles

wb2.worksheets[0].insert_row(0)

wb2.worksheets[0].change_column_width(3,50)

wb2.worksheets[0].insert_column(0)

p wb2.worksheets[0].row_styles
# raise 'end'

puts 'begin writing ' + file_path
# p wb2
wb2[0].change_row_font_color(2,'ffffff')

wb2.write(file_path)
puts 'completed writing ' + file_path

# app = Writer::AppWriter.new(dirpath, wb)
# p app.hash
# two = app.hash
# puts ''
# app.write()
# 
# 
# str = XmlSimple.xml_out(two)
# str = str.gsub(/opt/,'Properties')
# puts '..'
# p str


end