module RubyXL
  class Reference
    ROW_MAX = 1024*1024
    COL_MAX = 16393

    attr_reader :first_row, :last_row, :first_col, :last_col, :first_sheet, :last_sheet
    attr_accessor :base

    # RubyXL::Reference.new(row, col)
    # RubyXL::Reference.new(row_from, row_to, col_from, col_to)
    # RubyXL::Reference.new(reference_string)
    def initialize(*params)
      @first_sheet = @last_sheet = row_from = row_to = col_from = col_to = nil

      case params.size
      when 4 then row_from, row_to, col_from, col_to = params
      when 2 then row_from, col_from = params
      when 1 then
        raise ArgumentError.new("invalid value for #{self.class}: #{params[0].inspect}") unless params[0].is_a?(String)
        str = params[0]

        str =~ /\A(?:(?:(?:'([^':]+)(?:\:([^':]+))?)'|(?:(\w+)(?:\:(\w+))?))\!)?(\$?[A-Z]+\$?\d+)(?:\:(\$?[A-Z]+\$?\d+))?\Z/

        @first_sheet = $1 || $3
        @last_sheet  = $2 || $4 || @first_sheet
        row_from, col_from, row_from_abs, col_from_abs = self.class.ref2ind($5)
        row_to, col_to, row_to_abs, col_to_abs = self.class.ref2ind($6) unless $6.nil?
      end

      @first_row = row_from || 0
      @last_row  = row_to || row_from || ROW_MAX
      @first_col = col_from || 0
      @last_col  = col_to || col_from || COL_MAX
      @first_row_abs = row_from_abs
      @first_col_abs = col_from_abs
      @last_row_abs  = row_to_abs || row_from_abs
      @first_col_abs = col_to_abs || col_from_abs
      @base = nil
    end

    def single_cell?
      (@first_row == @last_row) && (@first_col == @last_col)
    end

    def ==(other)
      other &&
        (@first_row == other.first_row) && (@last_row == other.last_row) && 
        (@first_col == other.first_col) && (@last_col == other.last_col)
    end

    def cover?(other)
      other &&
        (@first_row <= other.first_row) && (@last_row >= other.last_row) &&
        (@first_col <= other.first_col) && (@last_col >= other.last_col)
    end

    def to_s
      if single_cell? then
        self.class.ind2ref(@first_row, @first_col, @first_row_abs, @first_col_abs)
      else
        self.class.ind2ref(@first_row, @first_col) + ':' + self.class.ind2ref(@last_row, @last_col)
      end
    end

    def inspect
      if single_cell? then
        "#<#{self.class} @row=#{@first_row} @col=#{@first_col}>"
      else
        "#<#{self.class} @row_range=#{@first_row..@last_row} @col_range=#{@col_range..@last_col}>"
      end
    end

    # Converts +row+ and +col+ zero-based indices to Excel-style cell reference
    # (0) A...Z, AA...AZ, BA... ...ZZ, AAA... ...AZZ, BAA... ...XFD (16383)
    def self.ind2ref(row = 0, col = 0, row_abs = false, col_abs = false)
      str = ''

      loop do
        x = col % 26
        str = ('A'.ord + x).chr + str
        col = (col / 26).floor - 1
        break if col < 0
      end

      str = ('$' + str) if col_abs
      str += '$' if row_abs

      str += (row + 1).to_s
    end

    # Converts Excel-style cell reference to +row+ and +col+ zero-based indices.
    def self.ref2ind(str)
      return [ -1, -1 ] unless str =~ /\A(\$)?([A-Z]+)(\$)?(\d+)\Z/

      col = 0
      $2.each_byte { |chr| col = col * 26 + (chr - 64) }
      [ $4.to_i - 1, col - 1, !$3.nil?, !$1.nil? ]
    end

  end

  class Sqref < Array

    def initialize(str)
      str.split.each { |ref_str| self << RubyXL::Reference.new(ref_str) }
    end

    def to_s
      self.collect{ |ref| ref.to_s }.join(' ')
    end

  end
end
