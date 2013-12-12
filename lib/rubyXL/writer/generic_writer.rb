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

  end
end
end
