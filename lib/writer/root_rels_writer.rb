# require File.expand_path(File.join(File.dirname(__FILE__),'workbook'))
# require File.expand_path(File.join(File.dirname(__FILE__),'worksheet'))
# require File.expand_path(File.join(File.dirname(__FILE__),'cell'))
# require File.expand_path(File.join(File.dirname(__FILE__),'color'))
require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class RootRelsWriter
    attr_accessor :dirpath, :workbook

    def initialize(dirpath, wb)
      @dirpath = dirpath
      @wb = wb
    end

    def write()
      contents = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>'
      contents
    end
  end
end
end
