module RubyXL
  class Color

    attr_accessor :theme, :indexed, :tint, :rgb, :auto

    def initialize(attrs = {})
      @theme        = attrs['theme']
      @indexed      = attrs['indexed']
      @tint         = attrs['tint']
      @rgb          = attrs['rgb']
      @auto         = attrs['auto']
    end

    #validates hex color code, no '#' allowed
    def self.validate_color(color)
      if color =~ /^([a-f]|[A-F]|[0-9]){6}$/
        return true
      else
        raise 'invalid color'
      end
    end

    def self.parse(node)
      color = self.new
      color.auto    = RubyXL::Parser.attr_int(node, 'auto')
      color.indexed = RubyXL::Parser.attr_int(node, 'indexed')
      color.theme   = RubyXL::Parser.attr_int(node, 'theme')
      color.tint    = RubyXL::Parser.attr_float(node, 'tint')
      color.rgb     = RubyXL::Parser.attr_string(node, 'rgb')
      color
    end

    def build_xml(xml, node_name = 'color')
      @attrs = {}
      @attrs[:auto]    = @auto    unless @auto.nil?
      @attrs[:indexed] = @indexed unless @indexed.nil?
      @attrs[:theme]   = @theme   unless @theme.nil?
      @attrs[:tint]    = @tint    unless @tint.nil?
      @attrs[:rgb]     = @rgb     unless @rgb.nil?
      xml.send(node_name, @attrs)
    end

  end

end
