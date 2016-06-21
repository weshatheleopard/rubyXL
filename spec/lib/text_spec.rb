require 'spec_helper'

describe RubyXL::Text do

  describe '.to_s' do

    it 'should not crash processing UNICODE data' do
      bytes = [ 114, 39, 95, 120, 48, 48, 56, 48, 95, 226, 132, 162, 115,
                32, 103, 105, 114, 108, 102, 114, 105, 101, 110, 100,
                39, 95, 120, 48, 48, 56, 48, 95, 226, 132, 162, 115, 32, 104, 111]

      t = RubyXL::Text.new(:value => bytes.pack("c*").force_encoding('UTF-8'))

      str = t.to_s

      expect(str).to be
    end

  end

end
