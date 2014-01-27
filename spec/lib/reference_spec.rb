require 'rubygems'
require 'rubyXL'

describe RubyXL::Reference do

  describe '.ind2ref + .ref2ind' do
    it 'should correctly return the "Excel Style" description of cells when given a row/column number' do
      RubyXL::Reference.ind2ref(0, 26).should == 'AA1'
      RubyXL::Reference.ind2ref(99, 0).should == 'A100'
      RubyXL::Reference.ind2ref(0, 26).should == 'AA1'
      RubyXL::Reference.ind2ref(0, 51).should == 'AZ1'
      RubyXL::Reference.ind2ref(0, 52).should == 'BA1'
      RubyXL::Reference.ind2ref(0, 77).should == 'BZ1'
      RubyXL::Reference.ind2ref(0, 78).should == 'CA1'
      RubyXL::Reference.ind2ref(0, 16383).should == 'XFD1'
    end

    it 'should correctly convert back and forth between "Excel Style" and index style cell references' do
      0.upto(16383) do |n|
        RubyXL::Reference.ref2ind(RubyXL::Reference.ind2ref(n, 16383 - n)).should == [ n, 16383 - n ]
      end
    end

    it 'should return [-1, -1] if the Excel index is not well-formed' do
      RubyXL::Reference.ref2ind('A1B').should == [-1, -1]
    end
  end

end