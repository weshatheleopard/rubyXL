module RubyXL
  class Color
    attr_accessor :theme, :indexed, :tint, :rgb

    #validates hex color code, no '#' allowed
    def self.validate_color(color)
      if color =~ /^([a-f]|[A-F]|[0-9]){6}$/
        return true
      else
        raise 'invalid color'
      end
    end

    def self.parse(xml)
      color = self.new

      indexed = xml.attributes['indexed']
      color.indexed = indexed && Integer(indexed.value)

      theme = xml.attributes['theme']
      color.theme = theme && Integer(theme.value)

      tint = xml.attributes['tint']
      color.tint = tint && Float(tint.value)

      rgb = xml.attributes['rgb']
      color.rgb = rgb && rgb.value

      color
    end

  end

end
