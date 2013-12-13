require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer
  class GenericWriter

    def initialize(workbook)
      @workbook = workbook
      # +self.class+ makes sure constant is pulled from descendant class, not from this one.
      # self.class::FILEPATH
    end

    def filepath
      raise 'Subclass responsebility'
    end

    def build_xml
      seed_xml = Nokogiri::XML('<?xml version = "1.0" standalone ="yes"?>')
      seed_xml.encoding = 'UTF-8'
  
      builder = Nokogiri::XML::Builder.with(seed_xml) do |param|
                  yield(param)
                end

      builder.to_xml({ :indent => 0, :save_with => Nokogiri::XML::Node::SaveOptions::AS_XML })
    end

    def add_to_zip(zipfile)
      zipfile.get_output_stream(filepath) { |f| f << write }
    end

  end
end
end
