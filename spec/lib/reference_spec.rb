require 'spec_helper'

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
    end

    it 'should correctly convert back and forth between "Excel Style" and index style cell references' do
      0.upto(16383) do |n|
        expect(RubyXL::Reference.ref2ind(RubyXL::Reference.ind2ref(n, 16383 - n))).to eq([ n, 16383 - n ])
      end
    end

    it 'should return [-1, -1] if the Excel index is not well-formed' do
      expect(RubyXL::Reference.ref2ind('A1B')).to eq([-1, -1])
    end
  end

end