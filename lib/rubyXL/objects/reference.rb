# frozen_string_literal: true

module RubyXL
  class Reference
    ROW_MAX = 1024 * 1024
    COL_MAX = 16393

    attr_reader :first_row, :last_row, :first_col, :last_col, :sheet_name
    attr_reader :row_from_absolute, :row_to_absolute, :col_from_absolute, :col_to_absolute

    # RubyXL::Reference.new(row, col)
    # RubyXL::Reference.new(row_from, row_to, col_from, col_to)
    # RubyXL::Reference.new(reference_string)
    # RubyXL::Reference.new(row_from:, row_to:, col_from:, col_to:)
    def initialize(*params)
      row_from = row_to = col_from = col_to = nil
      @row_from_absolute = @row_to_absolute = @col_from_absolute = @col_to_absolute = false

      case params.size
      when 4 then row_from, row_to, col_from, col_to = params
      when 2 then row_from, col_from = params
      when 1 then
        case params.first
        when Hash then
          row_from, row_to, col_from, col_to = params.first.fetch_values(:row_from, :row_to, :col_from, :col_to)
        when String then
          str = params.first
          match_data = str.match(/^('(?<sheet_name1>[^']+)'|(?<sheet_name2>[^']+))!/)
          if match_data then
            @sheet_name = match_data['sheet_name1'] || match_data['sheet_name2']
            str = str[match_data[0].size..-1]
          end

          from, to = str.split(':')
          row_from, col_from, @row_from_absolute, @col_from_absolute = self.class.ref2ind(from)
          row_to, col_to, @row_to_absolute, @col_to_absolute = self.class.ref2ind(to) unless to.nil?
        else
          raise ArgumentError.new("invalid value for #{self.class}: #{params[0].inspect}") unless params[0].is_a?(String)
        end
      end

      @first_row = row_from || 0
      @last_row = row_to || row_from || ROW_MAX
      @first_col = col_from || 0
      @last_col = col_to || col_from || COL_MAX
      @row_range = nil
      @col_range = nil
    end

    def row_range
      @row_range ||= Range.new(@first_row, @last_row)
    end

    def col_range
      @col_range ||= Range.new(@first_col, @last_col)
    end

    def single_cell?
      (@first_row == @last_row) && (@first_col == @last_col)
    end

    def valid?
      !(@first_row.negative? || @first_col.negative?)
    end

    def ==(other)
      return false if other.nil?
      return false unless @sheet_name == other.sheet_name

      if other.is_a?(RubyXL::Reference)
        (@first_row == other.first_row) && (@last_row == other.last_row) &&
          (@first_col == other.first_col) && (@last_col == other.last_col)
      else
        (row_range == other.row_range) && (col_range == other.col_range)
      end
    end

    def cover?(other)
      return false if other.nil?

      if other.is_a?(RubyXL::Reference)
        (@first_row <= other.first_row) && (other.last_row <= @last_row) &&
          (@first_col <= other.first_col) && (other.last_col <= @last_col)
      else
        row_range.cover?(other.row_range.begin) &&
          row_range.cover?(other.row_range.end) &&
          col_range.cover?(other.col_range.begin) &&
          col_range.cover?(other.col_range.end)
      end
    end

    def to_s
      result = +''

      if @sheet_name then
        if @sheet_name.index(' ') then
          result << "'#{@sheet_name}'"
        else
          result << @sheet_name
        end
        result << '!'
      end

      if single_cell? then
        result << self.class.ind2ref(@first_row, @first_col, @row_from_absolute, @col_from_absolute)
      else
        result << self.class.ind2ref(@first_row, @first_col, @row_from_absolute, @col_from_absolute)
        result << ':'
        result << self.class.ind2ref(@last_row, @last_col, @row_to_absolute, @col_to_absolute)
      end
    end

    def inspect
      if single_cell? then
        "#<#{self.class} @sheet_name=#{@sheet_name} @row=#{@first_row} @col=#{@first_col}>"
      else
        "#<#{self.class} @sheet_name=#{@sheet_name} @row_range=#{row_range} @col_range=#{col_range}>"
      end
    end

    # Converts +row+ and +col+ zero-based indices to Excel-style cell reference
    # <0> A...Z, AA...AZ, BA... ...ZZ, AAA... ...AZZ, BAA... ...XFD <16383>
    def self.ind2ref(row = 0, col = 0, row_abs = false, col_abs = false)
      col_ref = ''

      loop do
        x = col % 26
        col_ref = ('A'.ord + x).chr + col_ref
        col = (col / 26).floor - 1
        break if col < 0
      end

      "#{col_abs ? '$' : ''}#{col_ref}#{row_abs ? '$' : ''}#{row + 1}"
    end

    # Converts Excel-style cell reference to +row+ and +col+ zero-based indices.
    def self.ref2ind(str)
      matchdata = str.match(/\A(?<cabs>\$?)(?<col>[A-Z]+)(?<rabs>\$?)(?<row>\d+)\Z/)
      return [ -1, -1 ] unless matchdata
      [  matchdata['row'].to_i - 1,
         matchdata['col'].each_byte.inject(0) { |col, chr| (col * 26) + (chr - 64) } - 1,
         !matchdata['rabs'].empty?,
         !matchdata['cabs'].empty? ]
    end
  end

  class Sqref < Array
    def initialize(str)
      str.split.each { |ref_str| self << RubyXL::Reference.new(ref_str) }
    end

    def to_s
      self.collect(&:to_s).join(' ')
    end
  end
end
