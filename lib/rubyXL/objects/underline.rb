require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/simple_types'

module RubyXL
  class UnderlineValue < OOXMLObject
    # According to ST_UnderlineValues in the OOXML schema, a u element in a stylesheet
    # can have val="single", val="double", val="singleAccounting", val="doubleAccounting",
    # val="none" or no val attribute at all (same as val="single").
    define_attribute(:val, :string, :required => false)
  end
end
