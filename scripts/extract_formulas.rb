require 'rubyXL'

spreadsheets = Dir.glob(File.join("test", "input", "*.xlsx")).sort!

spreadsheets.each { |input|
  doc = RubyXL::Parser.parse(input)

  doc.worksheets.each { |ws|
    next unless ws.is_a? RubyXL::Worksheet
    ws.sheet_data.rows.each { |r|
      next if r.nil?
      r.cells.each {  | c|
        next if c.nil?
        f = c.formula
        next if f.nil? || f.t == 'shared'
        puts f.expression
      }
    }
  }
}   
