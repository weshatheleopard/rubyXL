require '../rubyXL'
require 'rubygems'
# require File.expand_path(File.join(File.dirname(__FILE__),'workbook'))
# require File.expand_path(File.join(File.dirname(__FILE__),'parser'))

module RubyXL

wb = Parser.parse('/Users/vbhagwat/Desktop/macros.xlsm', false)
wb.write('/Users/vbhagwat/Desktop/test2/Output/macros.xlsm')
puts "completed writing /Users/vbhagwat/Desktop/test2/Output/macros.xlsm"

end