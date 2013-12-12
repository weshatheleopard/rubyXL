require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class GenericWriter
    attr_reader :dirpath, :filepath, :workbook

    def initialize(dirpath, workbook)
      @dirpath = dirpath
      @workbook = workbook
      # +self.class+ makes sure constant is pulled from descendant class, not from this one.
      @filepath = File.join(dirpath, self.class::FILEPATH)
    end

    def build_xml
      seed_xml = Nokogiri::XML('<?xml version = "1.0" standalone ="yes"?>')
      seed_xml.encoding = 'UTF-8'
  
      builder = Nokogiri::XML::Builder.with(seed_xml) do |param|
                  yield(param)
                end

      builder.to_xml({ :indent => 0, :save_with => Nokogiri::XML::Node::SaveOptions::AS_XML })
    end

  end
end
end
