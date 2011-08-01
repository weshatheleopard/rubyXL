# require File.expand_path(File.join(File.dirname(__FILE__),'workbook'))
# require File.expand_path(File.join(File.dirname(__FILE__),'worksheet'))
# require File.expand_path(File.join(File.dirname(__FILE__),'cell'))
# require File.expand_path(File.join(File.dirname(__FILE__),'color'))
require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class AppWriter
    attr_accessor :dirpath, :filepath, :workbook

    def initialize(dirpath, wb)
      @dirpath = dirpath
      @filepath = dirpath+'/docProps/app.xml'
      @workbook = wb
    end

    def write()

      builder = Nokogiri::XML::Builder.new do |xml|
        xml.Properties('xmlns' => 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
        'xmlns:vt'=>'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes') {
          xml.Application @workbook.application
          xml.DocSecurity '0'
          xml.ScaleCrop 'false'
          xml.HeadingPairs {
            xml['vt'].vector(:size => '2', :baseType => 'variant') {
              xml['vt'].variant {
                xml['vt'].lpstr 'Worksheets'
              }
              xml['vt'].variant {
                xml['vt'].i4 @workbook.worksheets.size.to_s()
              }
            }
          }
          xml.TitlesOfParts {
            xml['vt'].vector(:size=>@workbook.worksheets.size.to_s(), :baseType=>"lpstr") {
              @workbook.worksheets.each do |sheet|
                xml['vt'].lpstr sheet.sheet_name
              end
            }
          }
          xml.Company @workbook.company
          xml.LinksUpToDate 'false'
          xml.SharedDoc 'false'
          xml.HyperlinksChanged 'false'
          xml.AppVersion @workbook.appversion
        }
      end
      contents = builder.to_xml
      if(contents =~ /xmlns:vt=\"(.*)\" xmlns=\"(.*)\"/)
        contents.sub(/xmlns:vt=\"(.*)\" xmlns=\"(.*)\"<A/,'xmlns="'+$2+'" xmlns:vt="'+$1+'"<A')
      end
      contents = contents.gsub(/\n/,'')
      contents = contents.gsub(/>(\s)+</,'><')
      contents = contents.sub(/<\?xml version=\"1.0\"\?>/,'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+"\n")
      contents
    end
  end
end
end
