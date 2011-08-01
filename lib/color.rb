module RubyXL
class Color

  #validates hex color code, no '#' allowed
  def Color.validate_color(color)
    if color =~ /^([a-f]|[A-F]|[0-9]){6}$/
      return true
    else
      raise 'invalid color'
    end
  end

end
end
