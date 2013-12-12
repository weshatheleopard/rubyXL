require 'rubygems'
require 'nokogiri'

module RubyXL
module Writer

  #TODO
  class CalcChainWriter < GenericWriter
    FILEPATH = '/xl/calcChain.xml'

    def write()
      contents = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'+"\n"
      contents += ''
      contents
    end
  end

end
end
