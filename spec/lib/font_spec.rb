require 'spec_helper'

describe RubyXL::Font do
  describe '.u' do
    it 'should preserve all valid values of enum DocumentFormat.OpenXml.Spreadsheet.UnderlineValues' do
      xml_template = '<font xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><u val="%s"/></font>'
      expect(RubyXL::Font.parse(xml_template % 'single').u.val).to eq('single')
      expect(RubyXL::Font.parse(xml_template % 'double').u.val).to eq('double')
      expect(RubyXL::Font.parse(xml_template % 'singleAccounting').u.val).to eq('singleAccounting')
      expect(RubyXL::Font.parse(xml_template % 'doubleAccounting').u.val).to eq('doubleAccounting')
      expect(RubyXL::Font.parse(xml_template % 'none').u.val).to eq('none')
    end

    it 'should interpret u element without val attribute as single underline' do
      xml_before = '<font xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><u/></font>'
      font = RubyXL::Font.parse(xml_before)
      xml_after = font.write_xml()
      expect(xml_after).to include('<u/>').or include('<u val="single"/>')
    end

    it 'should not add underlining automatically' do
      xml_before = '<font xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"></font>'
      font = RubyXL::Font.parse(xml_before)
      xml_after = font.write_xml()
      expect(xml_after).not_to include('<u')
    end
  end

end
