# require File.expand_path(File.join(File.dirname(__FILE__),'workbook'))
# require File.expand_path(File.join(File.dirname(__FILE__),'worksheet'))
# require File.expand_path(File.join(File.dirname(__FILE__),'cell'))
# require File.expand_path(File.join(File.dirname(__FILE__),'color'))
require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class WorkbookRelsWriter
    attr_accessor :dirpath, :filepath, :workbook

    def initialize(dirpath, wb)
      @dirpath = dirpath
      @workbook = wb
    end

    #all attributes out of order
    def write()
      builder = Nokogiri::XML::Builder.new do |xml|
        xml.Relationships('xmlns'=>'http://schemas.openxmlformats.org/package/2006/relationships') {
          i = 1
          @workbook.worksheets.each do |sheet|
            xml.Relationship('Id'=>'rId'+i.to_s,
              'Type'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet",
              'Target'=>'worksheets/sheet'+i.to_s+'.xml')
            i += 1
          end
          unless @workbook.external_links.nil?
            1.upto(@workbook.external_links.size-1) do |link|
              xml.Relationship('Id'=>'rId'+i.to_s,
                'Type'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
                'Target'=>"externalLinks/externalLink#{link}.xml"
              )
              i+=1
            end
          end
          xml.Relationship('Id'=>'rId'+i.to_s,
            'Type'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            'Target'=>'theme/theme1.xml')
          i += 1
          xml.Relationship('Id'=>'rId'+i.to_s,
            'Type'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
            'Target'=>'styles.xml')
          i+=1
          xml.Relationship('Id'=>'rId'+i.to_s,
            'Type'=>"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings",
              'Target'=>'sharedStrings.xml')
        }
      end
      contents = builder.to_xml
      contents = contents.gsub(/\n/,'')
      contents = contents.gsub(/>(\s)+</,'><')
      contents = contents.sub(/<\?xml version=\"1.0\"\?>/,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+"\n")
      contents
    end
  end
end
end
