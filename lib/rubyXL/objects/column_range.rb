require 'rubyXL/objects/ooxml_object'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_col-1.html
  class ColumnRange < OOXMLObject
    define_attribute(:min,          :uint,    :required => true)
    define_attribute(:max,          :uint,    :required => true)
    define_attribute(:width,        :double)
    define_attribute(:style,        :uint,    :default => 0, :accessor => :style_index)
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

    def include?(col_index)
      ((min-1)..(max-1)).include?(col_index)
    end

    DEFAULT_WIDTH = 9
  end

  class ColumnRanges < OOXMLObject
    define_child_node(RubyXL::ColumnRange, :collection => true, :accessor => :column_ranges)
    define_element_name 'cols'

    # Locate an existing column range, make a new one if not found,
    # or split existing column range into multiples.
    def get_range(col_index)
      col_num = col_index + 1

      old_range = self.find(col_index)

      if old_range.nil? then
        new_range = RubyXL::ColumnRange.new(:min => col_num, :max => col_num)
        self.column_ranges << new_range
        return new_range
      elsif old_range.min == col_num && 
              old_range.max == col_num then # Single column range, OK to change in place
        return old_range
      else
        raise "Range splitting not implemented yet"
      end
    end

    def find(col_index)
      column_ranges && column_ranges.find { |range| range.include?(col_index) }
    end

    def insert_column(col_index)
      column_ranges && column_ranges.each { |range| range.insert_column(col_index) }
    end

    def before_write_xml
      !(column_ranges.nil? || column_ranges.empty?)
    end

  end

end
