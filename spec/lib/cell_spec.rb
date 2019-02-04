require 'spec_helper'
require 'bigdecimal'
require 'rubyXL/convenience_methods/font'
require 'rubyXL/convenience_methods/cell'

describe RubyXL::Cell do

  before do
    @workbook  = RubyXL::Workbook.new
    @worksheet = @workbook.add_worksheet('Test Worksheet')
    @workbook.worksheets << @worksheet
    (0..10).each do |i|
      (0..10).each do |j|
        @worksheet.add_cell(i, j, "#{i}:#{j}")
      end
    end
    @cell = @worksheet[0][0]
  end

  describe '.add_cell' do
    it 'should properly assign data types' do
      r = 4
      c = 4

      cell = @worksheet.add_cell(r, c, 123)
      expect(cell.datatype).to be_nil

      cell = @worksheet.add_cell(r, c, "#{r}:#{c}")
      expect(cell.datatype).to eq(RubyXL::DataType::RAW_STRING)

      cell = @worksheet.add_cell(r, c, RubyXL::RichText.new(:t => RubyXL::Text.new(:value => 'Hello')))
      expect(cell.datatype).to eq(RubyXL::DataType::INLINE_STRING)
    end
  end

  describe '.change_fill' do
    it 'should cause an error if hex color code not passed' do
      expect {
        @cell.change_fill('G')
      }.to raise_error(RuntimeError)
    end

    it 'should make cell fill color equal to hex color code passed' do
      @cell.change_fill('0f0f0f')
      expect(@cell.fill_color).to eq('0f0f0f')
    end

    it 'should cause an error if hex color code includes # character' do
      expect {
        @cell.change_fill('#0f0f0f')
      }.to raise_error(RuntimeError)
    end
  end

  describe '.change_font_name' do
    it 'should make font name match font name passed' do
      @cell.change_font_name('Arial')
      expect(@cell.font_name).to eq('Arial')
    end
  end

  describe '.change_font_size' do
    it 'should make font size match number passed' do
      @cell.change_font_size(30)
      expect(@cell.font_size).to eq(30)
    end

    it 'should cause an error if a string passed' do
      expect {
        @cell.change_font_size('20')
      }.to raise_error(RuntimeError)
    end
  end

  describe '.change_font_color' do
    it 'should cause an error if hex color code not passed' do
      expect {
        @cell.change_font_color('G')
      }.to raise_error(RuntimeError)
    end

    it 'should make cell font color equal to hex color code passed' do
      @cell.change_font_color('0f0f0f')
      expect(@cell.font_color).to eq('0f0f0f')
    end

    it 'should cause an error if hex color code includes # character' do
      expect {
        @cell.change_font_color('#0f0f0f')
      }.to raise_error(RuntimeError)
    end
  end

  describe '.change_font_italics' do
    it 'should make cell font italicized when true is passed' do
      expect(@cell.is_italicized).to be_nil
      @cell.change_font_italics(true)
      expect(@cell.is_italicized).to eq(true)
    end
  end

  describe '.change_font_bold' do
    it 'should make cell font bolded when true is passed' do
      expect(@cell.is_bolded).to be_nil
      @cell.change_font_bold(true)
      expect(@cell.is_bolded).to eq(true)
    end
  end

  describe '.change_font_underline' do
    it 'should make cell font underlined when true is passed' do
      expect(@cell.is_underlined).to be_nil
      @cell.change_font_underline(true)
      expect(@cell.is_underlined).to eq(true)
    end
  end

  describe '.change_font_strikethrough' do
    it 'should make cell font struckthrough when true is passed' do
      expect(@cell.is_struckthrough).to be_nil
      @cell.change_font_strikethrough(true)
      expect(@cell.is_struckthrough).to eq(true)
    end
  end

  describe '.change_horizontal_alignment' do
    it 'should cause cell to horizontally align as specified by the passed in string' do
      expect(@cell.horizontal_alignment).to be_nil
      @cell.change_horizontal_alignment('center')
      expect(@cell.horizontal_alignment).to eq('center')
    end
  end

  describe '.change_vertical_alignment' do
    it 'should cause cell to vertically align as specified by the passed in string' do
      expect(@cell.vertical_alignment).to be_nil
      @cell.change_vertical_alignment('center')
      expect(@cell.vertical_alignment).to eq('center')
    end
  end

  describe '.change_wrap' do
    it 'should cause cell to wrap align as specified by the passed in value' do
      expect(@cell.text_wrap).to be_nil
      @cell.change_text_wrap(true)
      expect(@cell.text_wrap).to eq(true)
    end
  end

  describe '.change_text_indent' do
    it 'should cause the cell to have the corresponding text indent' do
      expect(@cell.text_indent).to be_nil
      @cell.change_text_indent(2)
      expect(@cell.text_indent).to eq(2)
    end

    it "should not cause other cells with the same style to have text indent" do
      another_cell = @worksheet[1][0]
      another_cell.style_index = @cell.style_index
      expect(another_cell.text_indent).to be_nil
      @cell.change_text_indent(2)
      expect(another_cell.text_indent).to be_nil
    end
  end

  describe '.change_border_color' do
    it 'should cause cell to have a colored top border' do
      expect(@cell.get_border_color(:top)).to be_nil
      @cell.change_border_color(:top, 'FF0000')
      expect(@cell.get_border_color(:top)).to eq('FF0000')
    end

    it 'should cause cell to have a colored bottom border' do
      expect(@cell.get_border_color(:bottom)).to be_nil
      @cell.change_border_color(:bottom, 'FF0000')
      expect(@cell.get_border_color(:bottom)).to eq('FF0000')
    end

    it 'should cause cell to have a colored left border' do
      expect(@cell.get_border_color(:left)).to be_nil
      @cell.change_border_color(:left, 'FF0000')
      expect(@cell.get_border_color(:left)).to eq('FF0000')
    end

    it 'should cause cell to have a colored right border' do
      expect(@cell.get_border_color(:right)).to be_nil
      @cell.change_border_color(:right, 'FF0000')
      expect(@cell.get_border_color(:right)).to eq('FF0000')
    end

    it 'should cause cell to have a colored diagonal border' do
      expect(@cell.get_border_color(:diagonal)).to be_nil
      @cell.change_border_color(:diagonal, 'FF0000')
      expect(@cell.get_border_color(:diagonal)).to eq('FF0000')
    end

    it 'is not overridden if the border style is set afterwards' do
      expect(@cell.get_border_color(:top)).to be_nil
      expect(@cell.get_border(:top)).to be_nil
      @cell.change_border_color(:top, 'FF0000')
      @cell.change_border(:top, 'thin')
      expect(@cell.get_border_color(:top)).to eq('FF0000')
      expect(@cell.get_border(:top)).to eq('thin')
    end
  end

  describe '.change_border' do
    it 'should cause cell to have border at top with specified weight' do
      expect(@cell.get_border(:top)).to be_nil
      @cell.change_border(:top, 'thin')
      expect(@cell.get_border(:top)).to eq('thin')
    end

    it 'should cause cell to have border at right with specified weight' do
      expect(@cell.get_border(:right)).to be_nil
      @cell.change_border(:right, 'thin')
      expect(@cell.get_border(:right)).to eq('thin')
    end

    it 'should cause cell to have border at left with specified weight' do
      expect(@cell.get_border(:left)).to be_nil
      @cell.change_border(:left, 'thin')
      expect(@cell.get_border(:left)).to eq('thin')
    end

    it 'should cause cell to have border at bottom with specified weight' do
      expect(@cell.get_border(:bottom)).to be_nil
      @cell.change_border(:bottom, 'thin')
      expect(@cell.get_border(:bottom)).to eq('thin')
    end

    it 'should cause cell to have border at diagonal with specified weight' do
      expect(@cell.get_border(:diagonal)).to be_nil
      @cell.change_border(:diagonal, 'thin')
      expect(@cell.get_border(:diagonal)).to eq('thin')
    end
  end

  describe '.value' do
    it 'should return the value of a date' do
      date = Date.parse('January 1, 2011')
      @cell.change_contents(date)
      expect(@cell).to receive(:is_date?).at_least(1).and_return(true)
      expect(@cell.value).to eq(date)
    end

    it 'should properly return value of inlineStr' do
      cell = @worksheet.add_cell(5, 5, RubyXL::RichText.new(:t => RubyXL::Text.new(:value => 'Hello')))
      expect(cell.value).to eq('Hello')
    end

    it "should properly handle numeric values" do
      @cell.datatype = nil
      @cell.raw_value = '1'
      expect(@cell.value).to eq(1)

      @cell.raw_value = '10000000'
      expect(@cell.value).to eq(10000000)

      @cell.raw_value = '-10'
      expect(@cell.value).to eq(-10)

      @cell.raw_value = '0'
      expect(@cell.value).to eq(0)

      @cell.raw_value = '0.001'
      expect(@cell.value).to eq(0.001)

      @cell.raw_value = '-0.00000000001'
      expect(@cell.value).to eq(-0.00000000001)

      @cell.raw_value = '1E5'
      expect(@cell.value).to eq(100000.0)

      @cell.raw_value = '1E0'
      expect(@cell.value).to eq(1.0)

      @cell.raw_value = '1E-5'
      expect(@cell.value).to eq(0.00001)

      @cell.raw_value = '-1E5'
      expect(@cell.value).to eq(-100000.0)

      @cell.raw_value = '-1E0'
      expect(@cell.value).to eq(-1.0)

      @cell.raw_value = '-1E-5'
      expect(@cell.value).to eq(-0.00001)

      @cell.raw_value = '1DE-5'
      expect(@cell.value).to eq('1DE-5')
    end

    context '1900-based dates' do
      before(:each) { @workbook.date1904 = false }
      it 'should convert date numbers correctly' do
        date = 41019
        @cell.change_contents(date)
        expect(@cell).to receive(:is_date?).at_least(1).and_return(true)
        expect(@cell.value).to eq(Date.parse('April 20, 2012'))
        @cell.change_contents(35981)
        expect(@cell.value).to eq(Date.parse('July 5, 1998'))
        @cell.change_contents(0.019467592592592595)
        expect(@cell.value).to eq(DateTime.parse('1899-12-31T00:28:02+00:00'))
        @cell.change_contents(1)
        expect(@cell.value).to eq(Date.parse('January 1, 1900'))
        @cell.change_contents(59)
        expect(@cell.value).to eq(Date.parse('February 28, 1900'))
        @cell.change_contents(60)
        # There's no way Ruby can return the nonexistent February 29th, so it has to be March 1st:
        expect(@cell.value).to eq(Date.parse('March 1, 1900'))
        @cell.change_contents(61)
        expect(@cell.value).to eq(Date.parse('March 1, 1900'))
      end
    end

    context '1904-based dates' do
      before(:each) { @workbook.date1904 = true }
      it 'should convert date numbers correctly' do
        date = 39557
        @cell.change_contents(date)
        expect(@cell).to receive(:is_date?).at_least(1).and_return(true)
        expect(@cell.value).to eq(Date.parse('April 20, 2012'))
        @cell.change_contents(34519)
        expect(@cell.value).to eq(Date.parse('July 5, 1998'))
      end
    end

    context 'date before January 1, 1900' do
      it 'should parse as date' do
        @cell.set_number_format('h:mm:ss')
        @cell.datatype = nil

        @cell.raw_value = '0.97726851851851848'
        expect(@cell.is_date?).to be(true)
        expect(@cell.value).to eq(DateTime.parse('1899-12-31 23:27:16'))
        @cell.raw_value = '1.9467592592592595E-2'
        expect(@cell.is_date?).to be(true)
        expect(@cell.value).to eq(DateTime.parse('1899-12-31 00:28:02'))
      end
    end

  end

  describe '.change_contents' do
    it 'should cause cell value to match string or number that is passed in' do
      @cell.change_contents('TEST')
      expect(@cell.value).to eq('TEST')
      expect(@cell.formula).to be_nil
    end

    it 'should cause cell value to match a date that is passed in' do
      date = Date.parse('January 1, 2011')
      @cell.change_contents(date)
      expect(@cell).to receive(:is_date?).at_least(1).and_return(true)
      expect(@cell.value).to eq(date)
      expect(@cell.datatype).to be_nil
      expect(@cell.formula).to be_nil
    end

    it 'should case cell value to match a Float that is passed in' do
      number = 1.25
      @cell.change_contents(number)
      expect(@cell.value).to eq(number)
      expect(@cell.datatype).to be_nil
      expect(@cell.formula).to be_nil
    end

    it 'should case cell value to match an Integer that is passed in' do
      number = 1234567
      @cell.change_contents(number)
      expect(@cell.value).to eq(number)
      expect(@cell.datatype).to be_nil
      expect(@cell.formula).to be_nil
    end

    it 'should case cell value to match an BigDecimal that is passed in' do
      number = BigDecimal.new('1234.5678')
      @cell.change_contents(number)
      expect(@cell.value).to eq(number)
      expect(@cell.datatype).to be_nil
      expect(@cell.formula).to be_nil
    end

    it 'should cause cell value and formula to match what is passed in' do
      @cell.change_contents(nil, 'SUM(A2:A4)')
      expect(@cell.value).to be_nil
      expect(@cell.formula.expression).to eq('SUM(A2:A4)')
    end
  end

  describe '.is_italicized' do
    it 'should correctly return whether or not the cell\'s font is italicized' do
      @cell.change_font_italics(true)
      expect(@cell.is_italicized).to eq(true)
    end
  end

  describe '.is_bolded' do
    it 'should correctly return whether or not the cell\'s font is bolded' do
      @cell.change_font_bold(true)
      expect(@cell.is_bolded).to eq(true)
    end
  end

  describe '.is_underlined' do
    it 'should correctly return whether or not the cell\'s font is underlined' do
      @cell.change_font_underline(true)
      expect(@cell.is_underlined).to eq(true)
    end
  end

  describe '.is_struckthrough' do
    it 'should correctly return whether or not the cell\'s font is struckthrough' do
      @cell.change_font_strikethrough(true)
      expect(@cell.is_struckthrough).to eq(true)
    end
  end

  describe '.font_name' do
    it 'should correctly return the name of the cell\'s font' do
      @cell.change_font_name('Verdana')
      expect(@cell.font_name).to eq('Verdana')
    end
  end

  describe '.font_size' do
    it 'should correctly return the size of the cell\'s font' do
      @cell.change_font_size(20)
      expect(@cell.font_size).to eq(20)
    end
  end

  describe '.font_color' do
    it 'should correctly return the color of the cell\'s font' do
      @cell.change_font_color('0f0f0f')
      expect(@cell.font_color).to eq('0f0f0f')
    end

    it 'should return 000000 (black) if no font color has been specified for this cell' do
      expect(@cell.font_color).to eq('000000')
    end
  end

  describe '.fill_color' do
    it 'should correctly return the color of the cell\'s fill' do
      @cell.change_fill('000000')
      expect(@cell.fill_color).to eq('000000')
    end

    it 'should return ffffff (white) if no fill color has been specified for this cell' do
      expect(@cell.fill_color).to eq('ffffff')
    end
  end

  describe '.horizontal_alignment' do
    it 'should correctly return the type of horizontal alignment of this cell' do
      @cell.change_horizontal_alignment('center')
      expect(@cell.horizontal_alignment).to eq('center')
    end

    it 'should return nil if no horizontal alignment has been specified for this cell' do
      expect(@cell.horizontal_alignment).to be_nil
    end

    it 'should not create new XFs when changing alignment to already existing values' do
      @cell.change_horizontal_alignment('left')
      style_xf1 = @cell.style_index
      @cell.change_horizontal_alignment('right')
      expect(@cell.style_index).not_to eq(style_xf1)
      style_xf2 = @cell.style_index
      @cell.change_horizontal_alignment('left')
      expect(@cell.style_index).to eq(style_xf1)
      @cell.change_horizontal_alignment('right')
      expect(@cell.style_index).to eq(style_xf2)
    end
  end

  describe '.vertical_alignment' do
    it 'should correctly return the type of vertical alignment of this cell' do
      @cell.change_vertical_alignment('center')
      expect(@cell.vertical_alignment).to eq('center')
    end

    it 'should return nil if no vertical alignment has been specified for this cell' do
      expect(@cell.vertical_alignment).to be_nil
    end
  end

  describe '.border_top' do
    it 'should correctly return the weight of the border on top for this cell' do
      @cell.change_border(:top, 'thin')
      expect(@cell.get_border(:top)).to eq('thin')
    end

    it 'should return nil if no top border has been specified for this cell' do
      expect(@cell.get_border(:top)).to be_nil
    end
  end

  describe '.border_left' do
    it 'should correctly return the weight of the border on left for this cell' do
      @cell.change_border(:left, 'thin')
      expect(@cell.get_border(:left)).to eq('thin')
    end

    it 'should return nil if no left border has been specified for this cell' do
      expect(@cell.get_border(:left)).to be_nil
    end
  end

  describe '.border_right' do
    it 'should correctly return the weight of the border on right for this cell' do
      @cell.change_border(:right, 'thin')
      expect(@cell.get_border(:right)).to eq('thin')
    end

    it 'should return nil if no right border has been specified for this cell' do
      expect(@cell.get_border(:right)).to be_nil
    end
  end

  describe '.border_bottom' do
    it 'should correctly return the weight of the border on bottom for this cell' do
      @cell.change_border(:bottom, 'thin')
      expect(@cell.get_border(:bottom)).to eq('thin')
    end

    it 'should return nil if no bottom border has been specified for this cell' do
      expect(@cell.get_border(:bottom)).to be_nil
    end
  end

  describe '.border_diagonal' do
    it 'should correctly return the weight of the diagonal border for this cell' do
      @cell.change_border(:diagonal, 'thin')
      expect(@cell.get_border(:diagonal)).to eq('thin')
    end

    it 'should return nil if no diagonal border has been specified for this cell' do
      expect(@cell.get_border(:diagonal)).to be_nil
    end
  end

  describe '.text_rotation' do
    it 'should correctly return the rotation for this cell' do
      expect(@cell.text_rotation).to be_nil
      @cell.change_text_rotation(45)
      expect(@cell.text_rotation).to eq(45)
    end
  end

end
