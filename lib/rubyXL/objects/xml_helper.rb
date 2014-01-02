module RubyXL
  module XMLhelper

    def attr_optional(k, v)
      @attrs[k] = v unless v.nil?
    end

  end
end