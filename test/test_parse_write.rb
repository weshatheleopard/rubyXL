require 'rubyXL'
require 'benchmark'

spreadsheets = Dir.glob(File.join("test", "input", "*.xlsx")).sort!

spreadsheets.each { |input|
  print "<<<--- Parsing #{input}..."
  doc = nil
  tm = Benchmark.realtime { doc = RubyXL::Parser.parse(input) }
  puts "Elapsed: #{tm} sec"
  output = File.join("test", "output", File.basename(input))
  print "--->>> Writing #{output}..."
  tm = Benchmark.realtime { doc.write(output) }
  puts "Elapsed: #{tm} sec"
}
