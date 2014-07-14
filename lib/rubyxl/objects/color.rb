require 'rubyXL/objects/ooxml_object'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_color-4.html
  class Color < OOXMLObject
    define_attribute(:auto,    :bool)
    define_attribute(:indexed, :uint)
    define_attribute(:rgb,     :string)
    define_attribute(:theme,   :uint)
    define_attribute(:tint,    :double, :default => 0.0)
    define_element_name 'color'

    #validates hex color code, no '#' allowed
    def self.validate_color(color)
      if color =~ /\A([a-f]|[A-F]|[0-9]){6}\Z/
        return true
      else
        raise 'invalid color'
      end
    end

  end

end
