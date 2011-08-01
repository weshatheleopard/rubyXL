# require File.expand_path(File.join(File.dirname(__FILE__),'workbook'))
# require File.expand_path(File.join(File.dirname(__FILE__),'worksheet'))
# require File.expand_path(File.join(File.dirname(__FILE__),'cell'))
# require File.expand_path(File.join(File.dirname(__FILE__),'color'))
require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class CoreWriter
    attr_accessor :dirpath, :filepath, :workbook

    def initialize(dirpath, wb)
      @dirpath = dirpath
      @workbook = wb
      @filepath = @dirpath + '/docProps/core.xml'
    end

    def write()
      builder = Nokogiri::XML::Builder.new do |xml|
        xml.coreProperties('xmlns:cp'=>"http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
        'xmlns:dc'=>"http://purl.org/dc/elements/1.1/", 'xmlns:dcterms'=>"http://purl.org/dc/terms/",
        'xmlns:dcmitype'=>"http://purl.org/dc/dcmitype/", 'xmlns:xsi'=>"http://www.w3.org/2001/XMLSchema-instance") {
          xml['dc'].creator @workbook.creator.to_s
          xml['cp'].lastModifiedBy @workbook.modifier.to_s
          xml['dcterms'].created('xsi:type' => 'dcterms:W3CDTF') do
            @workbook.created_at
          end

          xml['dcterms'].modified('xsi:type' => 'dcterms:W3CDTF')
        }
      end

      contents = builder.to_xml
      contents = contents.gsub(/coreProperties/,'cp:coreProperties')
      contents = contents.gsub(/\n/,'')
      contents = contents.gsub(/>(\s)+</,'><')

      #seems hack-y..
      contents = contents.gsub(/<dcterms:created xsi:type=\"dcterms:W3CDTF\"\/>/,
        '<dcterms:created xsi:type="dcterms:W3CDTF">'+@workbook.created_at+'</dcterms:created>')
      contents = contents.gsub(/<dcterms:modified xsi:type=\"dcterms:W3CDTF\"\/>/,
        '<dcterms:modified xsi:type="dcterms:W3CDTF">'+@workbook.modified_at+'</dcterms:modified>')

      contents = contents.sub(/<\?xml version=\"1.0\"\?>/,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+"\n")

      return contents
    end
  end
end
end
