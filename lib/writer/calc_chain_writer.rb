# require File.expand_path(File.join(File.dirname(__FILE__),'workbook'))
# require File.expand_path(File.join(File.dirname(__FILE__),'worksheet'))
# require File.expand_path(File.join(File.dirname(__FILE__),'cell'))
# require File.expand_path(File.join(File.dirname(__FILE__),'color'))
require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer


   #TODO
  class CalcChainWriter
    attr_accessor :dirpath, :filepath, :workbook

    def initialize(dirpath, wb)
      @dirpath = dirpath
      @workbook = wb
      @filepath = dirpath + '/xl/calcChain.xml'
    end

    def write()
      contents = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+"\n"
      contents += ''
      # file = File.new(@filepath, 'w+')
      # file.write(contents)
      # file.close
      contents
    end
  end

end
end
