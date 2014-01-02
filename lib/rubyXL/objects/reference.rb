module RubyXL
  class Reference
    ROW_MAX = 1024*1024
    COL_MAX = 16393

    attr_reader :row_range, :col_range

    # RubyXL::Reference.new(row, col)
    # RubyXL::Reference.new(row_from, row_to, col_from, col_to)
    # RubyXL::Reference.new(reference_string)
    def initialize(*params)
      row_from = row_to = col_from = col_to = nil

      case params.size
      when 4 then row_from, row_to, col_from, col_to = params
      when 2 then row_from, col_from = params
      when 1 then
        raise ArgumentError.new("invalid value for #{self.class}: #{params[0].inspect}") unless params[0].is_a?(String)
        from, to = params[0].split(':')
        row_from, col_from = self.class.ref2ind(from)
        row_to, col_to = self.class.ref2ind(to) unless to.nil?
      end      

      @row_range = Range.new(row_from || 0, row_to || row_from || ROW_MAX)
      @col_range = Range.new(col_from || 0, col_to || col_from || COL_MAX)
    end

    def single_cell?
      (@row_range.begin == @row_range.end) && (@col_range.begin == @col_range.end)
    end

    def first_row
      @row_range.begin
    end

    def last_row
      @row_range.end
    end

    def first_col
      @col_range.begin
    end

    def last_col
      @col_range.end
    end

    def ==(other)
      !other.nil? && (@row_range == other.row_range) && (@col_range == other.col_range)
    end

    def cover?(other)
      !other.nil? && (@row_range.cover?(other.row_range.begin) && @row_range.cover?(other.row_range.end) && 
                      @col_range.cover?(other.col_range.begin) && @col_range.cover?(other.col_range.end))
    end

    def to_s
      if single_cell? then
        self.class.ind2ref(@row_range.begin, @col_range.begin)
      else
        self.class.ind2ref(@row_range.begin, @col_range.begin) + ':' + self.class.ind2ref(@row_range.end, @col_range.end)
      end
    end

    def inspect
      type = single_cell? ? 'SINGLE_CELL' : 'RANGE'
      "#<#{self.class} (#{type}) @row_range=#{@row_range} @col_range=#{@col_range}>"
    end

    # Converts +row+ and +col+ zero-based indices to Excel-style cell reference
    # (0) A...Z, AA...AZ, BA... ...ZZ, AAA... ...AZZ, BAA... ...XFD (16383)
    def self.ind2ref(row = 0, col = 0)
      RubyXL::ColumnRange.ind2ref(col) + (row + 1).to_s
    end

    # Converts Excel-style cell reference to +row+ and +col+ zero-based indices.
    def self.ref2ind(str)
      return [ -1, -1 ] unless str =~ /^([A-Z]+)(\d+)$/
      [ $2.to_i - 1, RubyXL::ColumnRange.ref2ind($1) ]
    end  

  end
end
