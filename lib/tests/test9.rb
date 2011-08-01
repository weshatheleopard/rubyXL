require 'rubygems'
require 'nokogiri'
require File.expand_path(File.join(File.dirname(__FILE__),'Hash'))

module RubyXL
testWb = Nokogiri::XML.parse(File.read('/Users/vbhagwat/Desktop/testWorkbook.xml'))
testStyles = Nokogiri::XML.parse(File.read('/Users/vbhagwat/Desktop/testStyles.xml'))

puts "testWb.css('workbook definedNames')[0]"
p testWb.css('workbook definedNames')[0]
puts '/css'




Hash.xml_node_to_hash(test_wb.css('workbook definedNames')[0])

puts '



















'
h= Hash.xml_node_to_hash(test_styles.css('styleSheet cellXfs')[0])

puts '





'
p h
# 
# :cellXfs=>{
#   :xf=>[{
#     :attributes=>{
#       :xfId=>0, :fontId=>0, :numFmtId=>0, :borderId=>0, :fillId=>0}}, 
#     {:attributes=>{
#       :applyFill=>0, :applyFont=>0, :applyAlignment=>1, :xfId=>0, :fontId=>0, 
#       :numFmtId=>0, :applyNumberFormat=>0, :borderId=>0, :applyBorder=>0, :fillId=>0}, 
#       :alignment=>{:attributes=>{:horizontal=>"center"}}}], 
#       
#   :attributes=>{:count=>2}}
end