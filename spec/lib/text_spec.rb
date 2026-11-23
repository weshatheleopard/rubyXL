# frozen_string_literal: true

require 'spec_helper'

describe RubyXL::Text do
  describe '.to_s' do
    it 'should not crash processing UNICODE data' do
      bytes = [ 114, 39, 95, 120, 48, 48, 56, 48, 95, 226, 132, 162, 115,
                32, 103, 105, 114, 108, 102, 114, 105, 101, 110, 100,
                39, 95, 120, 48, 48, 56, 48, 95, 226, 132, 162, 115, 32, 104, 111]

      t = RubyXL::Text.new(:value => bytes.pack('c*').force_encoding('UTF-8'))

      str = t.to_s

      expect(str).to be
    end

    it 'should not escape valid XML extended UNICODE characters' do
      string_with_valid_chars = "\u{10000}\u{10FABC}\u{10FFFF}"
      t = RubyXL::Text.new(:value => string_with_valid_chars.dup)
      xml = t.write_xml
      expect(xml[%r{<t>[^<]+</t>}]).to eq("<t>#{string_with_valid_chars}</t>")
    end

    it 'should escape invalid in XML characters' do
      string_with_invalid_chars = "VT - \v, BELL - \a, UNICODE BLOCK - [\u0001...\uFFFE]"
      t = RubyXL::Text.new(:value => string_with_invalid_chars.dup)
      xml = t.write_xml
      expect(xml[%r{<t>[^<]+</t>}]).to eq("<t>VT - _x000B_, BELL - _x0007_, UNICODE BLOCK - [_x0001_..._xFFFE_]</t>")
      expect(t.to_s).to eq(string_with_invalid_chars)
    end
  end
end

describe RubyXL::RichText do
  describe '.to_s' do
    it 'should concatenate rich text runs' do
      runs = [ 'text run 1', 'text run 2', 'text run 3']

      rt = RubyXL::RichText.new
      runs.each { |rtr|
        rt.r << RubyXL::RichTextRun.new(:t => RubyXL::Text.new(:value => rtr))
      }

      expect(rt.to_s).to eq(runs.join)
    end

    it 'in the absence of text runs it should use plaintext' do
      str = 'test string'

      rt = RubyXL::RichText.new
      rt.t = RubyXL::Text.new(:value => str)

      expect(rt.to_s).to eq(str)
    end
  end

end
