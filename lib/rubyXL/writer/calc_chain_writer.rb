require 'rubygems'
require 'nokogiri'

module RubyXL
  module Writer

    #TODO
    class CalcChainWriter < GenericWriter
      def filepath
        File.join('xl', 'calcChain.xml')
      end

      def write()
        build_xml do |xml|
          nil
        end
      end
    end

  end
end
