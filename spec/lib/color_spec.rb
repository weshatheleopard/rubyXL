require 'rubygems'
require 'rubyXL'

describe RubyXL::Color do
  describe '.validate_color' do
    it 'should return true if a valid hex color without a # is passed' do
      RubyXL::Color.validate_color('0fbCAd').should == true
    end

    it 'should cause an error if an invalid hex color code or one with a # is passed' do
      lambda {RubyXL::Color.validate_color('#G')}.should raise_error
    end
  end
end
