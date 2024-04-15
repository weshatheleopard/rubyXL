require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/simple_types'

module RubyXL
  # http://www.datypic.com/sc/ooxml/e-ssml_color-4.html
  class Color < OOXMLObject
    COLOR_REGEXP = /\A([a-f]|[A-F]|[0-9]){6}\Z/

    define_attribute(:auto,    :bool)
    define_attribute(:indexed, :uint)
    define_attribute(:rgb,     RubyXL::ST_UnsignedIntHex)
    define_attribute(:theme,   :uint)
    define_attribute(:tint,    :double, :default => 0.0)
    define_element_name 'color'

    # validates hex color code, no '#' allowed
    def self.validate_color(color)
      return true if color =~ COLOR_REGEXP

      raise 'invalid color'
    end
  end
end
