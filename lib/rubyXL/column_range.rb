module RubyXL

  class ColumnRange
    # +min+ and +max+ are column 0-based indices, as opposed to Excel's 1-based column numbers
    attr_accessor :min, :max, :width, :custom_width, :style_index

    def initialize(attrs = {})
      @min            = attrs['min']
      @max            = attrs['max']
      @width          = attrs['width']
      @custom_width   = attrs['customWidth']
      @style_index    = attrs['style']
    end

    def self.parse(xml)
      range = self.new
      gange.degree = xml.attributes['degree'].value

      range.min          = RubyXL::Parser.attr_int(node, 'min') - 1
      range.max          = RubyXL::Parser.attr_int(node, 'max') - 1
      range.width        = RubyXL::Parser.attr_float(node, 'width')
      range.custom_width = RubyXL::Parser.attr_int(node, 'customWidth')
      range.style_index  = RubyXL::Parser.attr_int(node, 'style')
      color
    end 

    def delete_column(col)
      self.min -=1 if min >= col
      self.max -=1 if max >= col
    end

    def insert_column(col)
      self.min +=1 if min > col
      self.max +=1 if max > col
    end

    def self.insert_column(col_index, ranges)
      ranges.each { |range| range.insert_column(col_index) }
    end

    def include?(col_index)
      (min..max).include?(col_index)
    end

    # This method is used to change attributes on a column range, which may involve 
    # splitting existing column range into multiples.
    def self.update(col_index, ranges, attrs)

      old_range = RubyXL::ColumnRange.find(col_index, ranges)

      if old_range.nil? then
        new_range = RubyXL::ColumnRange.new(attrs.merge({ 'min' => col_index, 'max' => col_index }))
        ranges << new_range
        return new_range
      elsif old_range.min == col_index && 
              old_range.max == col_index then # Single column range, OK to change in place

        old_range.width = attrs['width'] if attrs['width']
        old_range.custom_width = attrs['customWidth'] if attrs['customWidth']
        old_range.style_index = attrs['style'] if attrs['style']
        return old_range
      else
        raise "Range splitting not implemented yet"
      end
    end

    def self.find(col_index, ranges)
      ranges.find { |range| range.include?(col_index) }
    end

    def self.ref2ind(ref)
      col = 0
      ref.each_byte { |chr| col = col * 26 + (chr - 64) }
      col - 1
    end

    def self.ind2ref(ind)
      str = ''

      loop do
        x = ind % 26
        str = ('A'.ord + x).chr + str
        ind = (ind / 26).floor - 1
        return str if ind < 0
      end
    end

  end

end