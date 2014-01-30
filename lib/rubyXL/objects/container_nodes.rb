require 'rubyXL/objects/ooxml_object'

module RubyXL

  class BooleanValue < OOXMLObject
    define_attribute(:val, :bool, :required => true, :default => true)
  end

  class StringValue < OOXMLObject
    define_attribute(:val, :string, :required => true)
  end

  class IntegerValue < OOXMLObject
    define_attribute(:val, :int, :required => true)
  end

  class FloatValue < OOXMLObject
    define_attribute(:val, :float, :required => true)
  end

end