module RubyXL
  class SharedStrings

    attr_accessor :count_attr, :unique_count_attr

    def initialize
      # So far, going by the structure that the original creator had in mind. However, 
      # since the actual implementation is now extracted into a separate class, 
      # we will be able to transparrently change it later if needs be.
      @content_by_index = []
      @index_by_content = {}
      @unique_count_attr = @count_attr = nil # TODO
    end

    def empty?
      @content_by_index.empty?
    end

    def add(str, index)
      @content_by_index[index] = str
      @index_by_content[str] = index
    end

    def get_index(str, add_if_missing = false)
      index = @index_by_content[str]
      index = add(str) if index.nil? && add_if_missing
      index 
    end

    def[](index)
      @content_by_index[index]
    end

  end
end
