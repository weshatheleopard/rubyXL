module RubyXL

  class Font
    attr_accessor :count, :size, :name, :family, :color, :scheme, :bold, :italic, :underlined, :strikethrough

    def initialize
      @count = 0
      @size = @name = @family = @color = @scheme = nil
      @bold = @italic = @underlined = @strikethrough = false
    end

    def dup
      new = super
      new.count = 1
      new
    end

    def self.parse(xml)
      font = self.new
      xml.element_children.each { |node|
        case node.name
        when 'sz'     then font.size = RubyXL::Parser.attr_int(node, 'val')
        when 'name'   then font.name = RubyXL::Parser.attr_string(node, 'val')
        when 'family' then font.family = RubyXL::Parser.attr_int(node, 'val')
        when 'color'  then font.color = RubyXL::Color.parse(node)
        when 'scheme' then font.scheme = RubyXL::Parser.attr_string(node, 'val')
        when 'b'      then font.bold = true
        when 'i'      then font.italic = true
        when 'u'      then font.underlined = true
        when 'strike' then font.strikethrough = true
        end
      }
      font
    end 

    def build_xml(xml)
      xml.font {
        xml.sz(:val => @size)
        xml.b if @bold
        xml.i if @italic
        xml.u if @underlined
        xml.strike if @strikethrough
        @color.build_xml(xml) if @color
        xml.family(:val => @family) if @family
        xml.scheme(:val => @scheme) if @scheme
        xml.name(:val => @name) if @name
      }
    end

  end

end
