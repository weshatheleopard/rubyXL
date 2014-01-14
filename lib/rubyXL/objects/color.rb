module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_color-4.html
  class Color < OOXMLObject
    define_attribute(:auto,    :bool)
    define_attribute(:indexed, :int)
    define_attribute(:rgb,     :string)
    define_attribute(:theme,   :int)
    define_attribute(:tint,    :float)
    define_element_name 'color'

    #validates hex color code, no '#' allowed
    def self.validate_color(color)
      if color =~ /^([a-f]|[A-F]|[0-9]){6}$/
        return true
      else
        raise 'invalid color'
      end
    end

  end

end
