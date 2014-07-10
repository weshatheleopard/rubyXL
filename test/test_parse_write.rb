require 'rubyXL'
require 'benchmark'
require 'stackprof'

spreadsheets = Dir.glob(File.join("test", "input", "*.xls?")).sort!

spreadsheets.each { |input|
  puts "<<<--- Parsing #{input}..."
  doc = nil
#  tm = Benchmark.realtime { doc = RubyXL::Parser.parse(input) }
#  puts "Elapsed: #{tm} sec"
  StackProf.run(mode: :cpu, interval: 100, out: "stackprof-cpu-read-#{File.basename(input)}.dump") { doc = RubyXL::Parser.parse(input) }

#  output = File.join("test", "output", File.basename(input))
#  puts  "--->>> Writing #{output}..."
#  StackProf.run(mode: :cpu, out: "stackprof-cpu-write-#{File.basename(input)}.dump") { doc.write(output) }
#  tm = Benchmark.realtime 
#  puts "Elapsed: #{tm} sec"
}
