require 'rubygems'
require 'FileUtils'
require 'writer'


module RubyXL
  
def convertToIndex(cellString)
  index = Array.new(2)
  index[0]=-1
  index[1]=-1
  if(cellString =~ /^([A-Z]+)(\d+)/)
    one = $1.to_s()
    row = Integer($2) - 1 #-1 for 0 indexing
    col = 0
    i = 0
    one = one.reverse #because of 26^i calculation
    one.each_byte do |c|
      intVal = c - 64 #converts A to 1 (0, actually)
      col += intVal * 26**(i)
      i=i+1
    end
    col -= 1 #zer0 index
    index[0] = row
    index[1] = col
  end
  index
end


#TODO FIX THIS
def convertToCell(row, col)
  puts "input = (#{row}, #{col})"
  #convert col to letter, convert row to row-1, join
  rowString = (row + 1).to_s #+1 for 0 indexing
  colString = ''
  unless col == 0
    colLength = 1+Integer(Math.log(col) / Math.log(26)) #opposite of 26**
  else
    colLength = 1
  end

  puts "colLength = #{colLength}"
  1.upto(colLength) do |i|
    puts "i = #{i}"
    puts "col = #{col}"
    puts "26**(colLength-i) = #{26**(colLength-i)}"

    if i == colLength
      col+=1
    end

    if col >= 26**(colLength-i)
      intVal = col / 26**(colLength-i) #+1 for 0 indexing
      intVal += 64 #converts 1 to A, etc.
      puts "intVal = #{intVal}"
      puts "col = #{col}, #{intVal.chr}"
      
      # colString = intVal.chr + colString
      colString += intVal.chr
      
      puts "colString = #{colString}"
      col -= (intVal-64)*26**(colLength-i)
    end
  end
  p colString
  p rowString

  colString+rowString
end
row = 0
col = 16168

puts "(#{row},#{col})=>#{convertToCell(row,col)}"

end