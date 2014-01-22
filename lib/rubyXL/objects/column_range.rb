module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_col-1.html
  class ColumnRange < OOXMLObject
    define_attribute(:min,          :int,     :required => true)
    define_attribute(:max,          :int,     :required => true)
    define_attribute(:width,        :float)
    define_attribute(:style,        :int,     :default => 0)
    define_attribute(:hidden,       :bool,    :default => false)
    define_attribute(:bestFit,      :bool,    :default => false)
    define_attribute(:customWidth,  :bool,    :default => false)
    define_attribute(:phonetic,     :bool,    :default => false)
    define_attribute(:outlineLevel, :int,     :default => 0)
    define_attribute(:collapsed,    :bool,    :default => false)
    define_element_name 'col'

    def delete_column(col_index)
      col = col_index + 1
      self.min -=1 if min >= col
      self.max -=1 if max >= col
    end

    def insert_column(col_index)
      col = col_index + 1
      self.min +=1 if min >= col
      self.max +=1 if max >= col - 1
    end

    def self.insert_column(col_index, ranges)
      ranges.each { |range| range.insert_column(col_index) }
    end

    def include?(col_index)
      ((min-1)..(max-1)).include?(col_index)
    end

    # This method is used to change attributes on a column range, which may involve 
    # splitting existing column range into multiples.
    def self.update(col_index, ranges, attrs)
      col_num = col_index + 1

      old_range = RubyXL::ColumnRange.find(col_index, ranges)

      if old_range.nil? then
        new_range = RubyXL::ColumnRange.new(attrs.merge({ :min => col_num, :max => col_num }))
        ranges << new_range
        return new_range
      elsif old_range.min == col_num && 
              old_range.max == col_num then # Single column range, OK to change in place
        attrs.each_pair { |k, v| old_range.send("#{k}=", v) }
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

  class ColumnRanges < OOXMLObject
    define_child_node(RubyXL::ColumnRange, :collection => true, :accessor => :column_ranges)
    define_element_name 'cols'
  end

end
