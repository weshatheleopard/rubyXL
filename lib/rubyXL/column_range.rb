module RubyXL

  class ColumnRange
    # +min+ and +max+ are column 0-based indices, as opposed to Excel's 1-based column numbers
    attr_accessor :min, :max, :width, :custom_width, :style_index

    def initialize(attrs = {})
      @min            = get_attribute(attrs, 'min')
      @max            = get_attribute(attrs, 'max')
      @width          = get_attribute(attrs, 'width')
      @custom_width   = get_attribute(attrs, 'custom_width')
      @style_index    = get_attribute(attrs, 'style')
    end

    def update_attrs(attrs)
      @width          = get_attribute(attrs, 'width')        || @width
      @custom_width   = get_attribute(attrs, 'custom_width') || @custom_width
      @style_index    = get_attribute(attrs, 'style')        || @style_index
    end

    def get_attribute(attrs, k)
      v = attrs[k]
      v = v.value if v.is_a?(Nokogiri::XML::Attr)
      case v
      when String then
        intval = v.to_i rescue nil

        case intval.to_s
        when '' then v     # The value was not numeric to begin with, so do not change it
        when v then intval # The value converted right back into itself, which means it was integer.
        else Float(v)      # It converted into an Integer fine, but it was not an integer, so must be float.
        end
      when NilClass then nil
      else v
      end
    end 
    private :get_attribute

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
        old_range.update_attrs(attrs)
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