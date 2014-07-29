require 'rubygems'
require 'rubyXL'

describe RubyXL::Reference do

  describe '.ind2ref + .ref2ind' do
    it 'should correctly return the "Excel Style" description of cells when given a row/column number' do
      expect(RubyXL::Reference.ind2ref(0, 26)).to eq('AA1')
      expect(RubyXL::Reference.ind2ref(99, 0)).to eq('A100')
      expect(RubyXL::Reference.ind2ref(0, 26)).to eq('AA1')
      expect(RubyXL::Reference.ind2ref(0, 51)).to eq('AZ1')
      expect(RubyXL::Reference.ind2ref(0, 52)).to eq('BA1')
      expect(RubyXL::Reference.ind2ref(0, 77)).to eq('BZ1')
      expect(RubyXL::Reference.ind2ref(0, 78)).to eq('CA1')
      expect(RubyXL::Reference.ind2ref(0, 16383)).to eq('XFD1')

      expect(RubyXL::Reference.ind2ref(0, 1, false, false)).to eq('B1')
      expect(RubyXL::Reference.ind2ref(0, 1, true,  false)).to eq('B$1')
      expect(RubyXL::Reference.ind2ref(0, 1, false, true )).to eq('$B1')
      expect(RubyXL::Reference.ind2ref(0, 1, true,  true )).to eq('$B$1')
    end

    it 'should correctly convert back and forth between "Excel Style" and index style cell references' do
      0.upto(16383) do |n|
        expect(RubyXL::Reference.ref2ind(RubyXL::Reference.ind2ref(n, 16383 - n, false, false))).to eq([ n, 16383 - n, false, false ])
        expect(RubyXL::Reference.ref2ind(RubyXL::Reference.ind2ref(n, 16383 - n, true,  false))).to eq([ n, 16383 - n, true,  false ])
        expect(RubyXL::Reference.ref2ind(RubyXL::Reference.ind2ref(n, 16383 - n, false, true ))).to eq([ n, 16383 - n, false, true  ])
        expect(RubyXL::Reference.ref2ind(RubyXL::Reference.ind2ref(n, 16383 - n, true,  true ))).to eq([ n, 16383 - n, true,  true  ])
      end
    end

#    it 'should return [-1, -1] if the Excel index is not well-formed' do
#      expect(RubyXL::Reference.ref2ind('A1B')).to eq([-1, -1])
#    end

    it 'should properly parse cases from the manual' do
      # The cell in column A and row 10
      ref = RubyXL::Reference.new('A10')
      expect(ref.to_s).to eq('A10')
      expect(ref.single_cell?).to eq(true)
      expect(ref.first_row).to eq(9)
      expect(ref.first_col).to eq(0)
     

      # The range of cells in column A and rows 10 through 20
      ref = RubyXL::Reference.new('A10:A20')
      expect(ref.to_s).to eq('A10:A20')
      expect(ref.single_cell?).to eq(false)
      expect(ref.first_row).to eq(9)
      expect(ref.first_col).to eq(0)
      expect(ref.last_row).to eq(19)
      expect(ref.last_col).to eq(0)

      # The range of cells in row 15 and columns B through E
      ref = RubyXL::Reference.new('B15:E15')
      expect(ref.to_s).to eq('B15:E15')
      expect(ref.single_cell?).to eq(false)
      expect(ref.first_row).to eq(14)
      expect(ref.first_col).to eq(1)
      expect(ref.last_row).to eq(14)
      expect(ref.last_col).to eq(4)

      # All cells in row 5
      ref = RubyXL::Reference.new('5:5')
      expect(ref.to_s).to eq('5:5')
      expect(ref.single_cell?).to eq(false)
      expect(ref.first_row).to eq(4)
      expect(ref.first_col).to be_nil
      expect(ref.last_row).to eq(4)
      expect(ref.last_col).to be_nil

      # All cells in rows 5 through 10
      ref = RubyXL::Reference.new('5:10')
      expect(ref.to_s).to eq('5:10')
      expect(ref.single_cell?).to eq(false)
      expect(ref.first_row).to eq(4)
      expect(ref.first_col).to be_nil
      expect(ref.last_row).to eq(9)
      expect(ref.last_col).to be_nil

      # All cells in column H
      ref = RubyXL::Reference.new('H:H')
      expect(ref.to_s).to eq('H:H')
      expect(ref.single_cell?).to eq(false)
      expect(ref.first_row).to be_nil
      expect(ref.first_col).to eq(7)
      expect(ref.last_row).to be_nil
      expect(ref.last_col).to eq(7)

      # All cells in columns H through J
      ref = RubyXL::Reference.new('H:J')
      expect(ref.to_s).to eq('H:J')
      expect(ref.single_cell?).to eq(false)
      expect(ref.first_row).to be_nil
      expect(ref.first_col).to eq(7)
      expect(ref.last_row).to be_nil
      expect(ref.last_col).to eq(9)

      # The range of cells in row 15 and columns B through E
      ref = RubyXL::Reference.new('A10:E20')
      expect(ref.to_s).to eq('A10:E20')
      expect(ref.single_cell?).to eq(false)
      expect(ref.first_row).to eq(9)
      expect(ref.first_col).to eq(0)
      expect(ref.last_row).to eq(19)
      expect(ref.last_col).to eq(4)

    end

  end

end