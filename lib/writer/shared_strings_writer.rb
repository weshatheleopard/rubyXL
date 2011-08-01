# require File.expand_path(File.join(File.dirname(__FILE__),'workbook'))
# require File.expand_path(File.join(File.dirname(__FILE__),'worksheet'))
# require File.expand_path(File.join(File.dirname(__FILE__),'cell'))
# require File.expand_path(File.join(File.dirname(__FILE__),'color'))
require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
 class SharedStringsWriter
    attr_accessor :dirpath, :filepath, :workbook

    def initialize(dirpath,wb)
      @dirpath = dirpath
      @workbook = wb
      @filepath = dirpath + '/xl/sharedStrings.xml'
    end

    def write()
      # contents = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+"\n"

      # builder = Nokogiri::XML::Builder.new do |xml|
      #         xml.sst('xmlns'=>"http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      #         'count'=>@workbook.numStrings,
      #         'uniqueCount'=>@workbook.size) {
      #           i = 0
      #           0.upto(@workbook.size-1).each do |i|
      #             xml.si {
      #               xml.t @workbook.sharedStrings[i].to_s
      #               xml.phoneticPr('fontId'=>'1', 'type'=>'noConversion')
      #             }
      #           end
      #         }
      #       end
      #       contents = builder.to_xml
      #       contents = contents.gsub(/\n/,'')
      #       contents = contents.gsub(/>(\s)+</,'><')
      #       contents = contents.sub(/<\?xml version=\"1.0\"\?>/,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+"\n")
      contents = @workbook.shared_strings_XML
      contents
    end
  end
end
end
