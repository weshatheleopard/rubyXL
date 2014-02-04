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

      def render_xml
        seed_xml = Nokogiri::XML('<?xml version = "1.0" standalone ="yes"?>')
        seed_xml.encoding = 'UTF-8'

        yield(seed_xml)

        seed_xml.to_xml({ :indent => 0, :save_with => Nokogiri::XML::Node::SaveOptions::AS_XML })
      end

      def add_to_zip(zipfile)
        output = write
        return if output.nil?
        zipfile.get_output_stream(filepath) { |f| f << output }
      end

      def ooxml_object
        raise 'Subclass responsebility'
      end

      def write
        ooxml_object && ooxml_object.write_xml
      end

    end
  end
end
