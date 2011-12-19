require 'rubygems'
require 'rubyXL'

describe RubyXL::Cell do

  before do
    @workbook  = RubyXL::Workbook.new
    @worksheet = RubyXL::Worksheet.new(@workbook)
    @workbook.worksheets << @worksheet
    (0..10).each do |i|
      (0..10).each do |j|
        @worksheet.add_cell(i, j, "#{i}:#{j}")
      end
    end
    @cell = @worksheet[0][0]
  end

  describe '.change_fill' do
    it 'should cause an error if hex color code not passed' do
      lambda {
        @cell.change_fill('G')
      }.should raise_error
    end

    it 'should make cell fill color equal to hex color code passed' do
      @cell.change_fill('0f0f0f')
      @cell.fill_color.should == '0f0f0f'
    end

    it 'should cause an error if hex color code includes # character' do
      lambda {
        @cell.change_fill('#0f0f0f')
      }.should raise_error
    end
  end

  describe '.change_font_name' do
    it 'should make font name match font name passed' do
      @cell.change_font_name('Arial')
      @cell.font_name.should == 'Arial'
    end
  end

  describe '.change_font_size' do
    it 'should make font size match number passed' do
      @cell.change_font_size(30)
      @cell.font_size.should == 30
    end

    it 'should cause an error if a string passed' do
      lambda {
        @cell.change_font_size('20')
      }.should raise_error
    end
  end

  describe '.change_font_color' do
    it 'should cause an error if hex color code not passed' do
       lambda {
         @cell.change_font_color('G')
       }.should raise_error
     end

     it 'should make cell font color equal to hex color code passed' do
       @cell.change_font_color('0f0f0f')
       @cell.font_color.should == '0f0f0f'
     end

     it 'should cause an error if hex color code includes # character' do
       lambda {
         @cell.change_font_color('#0f0f0f')
       }.should raise_error
     end
  end

  describe '.change_font_italics' do
    it 'should make cell font italicized when true is passed' do
      @cell.change_font_italics(true)
      @cell.is_italicized.should == true
    end
  end

  describe '.change_font_bold' do
    it 'should make cell font bolded when true is passed' do
      @cell.change_font_bold(true)
      @cell.is_bolded.should == true
    end
  end

  describe '.change_font_underline' do
    it 'should make cell font underlined when true is passed' do
      @cell.change_font_underline(true)
      @cell.is_underlined.should == true
    end
  end

  describe '.change_font_strikethrough' do
    it 'should make cell font struckthrough when true is passed' do
      @cell.change_font_strikethrough(true)
      @cell.is_struckthrough.should == true
    end
  end

  describe '.change_horizontal_alignment' do
    it 'should cause cell to horizontally align as specified by the passed in string' do
       @cell.change_horizontal_alignment('center')
       @cell.horizontal_alignment.should == 'center'
     end

     it 'should cause error if nil, "center", "justify", "left", "right", or "distributed" is not passed' do
       lambda {
         @cell.change_horizontal_alignment('TEST')
       }.should raise_error
     end
  end

  describe '.change_vertical_alignment' do
    it 'should cause cell to vertically align as specified by the passed in string' do
       @cell.change_vertical_alignment('center')
       @cell.vertical_alignment.should == 'center'
     end

     it 'should cause error if nil, "center", "justify", "left", "right", or "distributed" is not passed' do
       lambda {
         @cell.change_vertical_alignment('TEST')
       }.should raise_error
     end
  end

  describe '.change_border_top' do
    it 'should cause cell to have border at top with specified weight' do
      @cell.change_border_top('thin')
      @cell.border_top.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @cell.change_border_top('TEST')
      }.should raise_error
    end
  end

  describe '.change_border_left' do
    it 'should cause cell to have border at left with specified weight' do
      @cell.change_border_left('thin')
      @cell.border_left.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @cell.change_border_left('TEST')
      }.should raise_error
    end
  end

  describe '.change_border_right' do
    it 'should cause cell to have border at right with specified weight' do
      @cell.change_border_right('thin')
      @cell.border_right.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @cell.change_border_right('TEST')
      }.should raise_error
    end
  end

  describe '.change_border_bottom' do
    it 'should cause cell to have border at bottom with specified weight' do
      @cell.change_border_bottom('thin')
      @cell.border_bottom.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @cell.change_border_bottom('TEST')
      }.should raise_error
    end
  end

  describe '.change_border_diagonal' do
    it 'should cause cell to have border at diagonal with specified weight' do
      @cell.change_border_diagonal('thin')
      @cell.border_diagonal.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @cell.change_border_diagonal('TEST')
      }.should raise_error
    end
  end

  describe '.value' do
    it 'should return the value of a date' do
      date = Date.parse('January 1, 2011')
      @cell.change_contents(date)
      @cell.should_receive(:is_date?).any_number_of_times.and_return(true)
      @cell.value.should == date
    end

    it 'should convert date numbers correctly' do
      date = 41019
      @cell.change_contents(date)
      @cell.should_receive(:is_date?).any_number_of_times.and_return(true)
      puts @cell.value
      puts Date.parse('April 20, 2012')
      @cell.value.should == Date.parse('April 20, 2012')
    end
  end

  describe '.change_contents' do
    it 'should cause cell value to match string or number that is passed in' do
      @cell.change_contents('TEST')
      @cell.value.should == 'TEST'
      @cell.formula.should == nil
    end

    it 'should cause cell value to match a date that is passed in' do
      date = Date.parse('January 1, 2011')
      @cell.change_contents(date)
      @cell.should_receive(:is_date?).any_number_of_times.and_return(true)
      @cell.value.should == date
      @cell.formula.should == nil
    end

    it 'should cause cell value and formula to match what is passed in' do
      @cell.change_contents(nil, 'SUM(A2:A4)')
      @cell.value.should == nil
      @cell.formula.should == 'SUM(A2:A4)'
    end
  end

  describe '.is_italicized' do
    it 'should correctly return whether or not the cell\'s font is italicized' do
      @cell.change_font_italics(true)
      @cell.is_italicized.should == true
    end
  end

  describe '.is_bolded' do
    it 'should correctly return whether or not the cell\'s font is bolded' do
      @cell.change_font_bold(true)
      @cell.is_bolded.should == true
    end
  end

  describe '.is_underlined' do
    it 'should correctly return whether or not the cell\'s font is underlined' do
      @cell.change_font_underline(true)
      @cell.is_underlined.should == true
    end
  end

  describe '.is_struckthrough' do
    it 'should correctly return whether or not the cell\'s font is struckthrough' do
      @cell.change_font_strikethrough(true)
      @cell.is_struckthrough.should == true
    end
  end

  describe '.font_name' do
    it 'should correctly return the name of the cell\'s font' do
      @cell.change_font_name('Verdana')
      @cell.font_name.should == 'Verdana'
    end
  end

  describe '.font_size' do
    it 'should correctly return the size of the cell\'s font' do
      @cell.change_font_size(20)
      @cell.font_size.should == 20
    end
  end

  describe '.font_color' do
    it 'should correctly return the color of the cell\'s font' do
      @cell.change_font_color('0f0f0f')
      @cell.font_color.should == '0f0f0f'
    end

    it 'should return 000000 (black) if no font color has been specified for this cell' do
      @cell.font_color.should == '000000'
    end
  end

  describe '.fill_color' do
    it 'should correctly return the color of the cell\'s fill' do
      @cell.change_fill('000000')
      @cell.fill_color.should == '000000'
    end

    it 'should return ffffff (white) if no fill color has been specified for this cell' do
      @cell.fill_color.should == 'ffffff'
    end
  end

  describe '.horizontal_alignment' do
    it 'should correctly return the type of horizontal alignment of this cell' do
      @cell.change_horizontal_alignment('center')
      @cell.horizontal_alignment.should == 'center'
    end

    it 'should return nil if no horizontal alignment has been specified for this cell' do
      @cell.horizontal_alignment.should == nil
    end
  end

  describe '.vertical_alignment' do
    it 'should correctly return the type of vertical alignment of this cell' do
      @cell.change_vertical_alignment('center')
      @cell.vertical_alignment.should == 'center'
    end

    it 'should return nil if no vertical alignment has been specified for this cell' do
      @cell.vertical_alignment.should be_nil
    end
  end

  describe '.border_top' do
    it 'should correctly return the weight of the border on top for this cell' do
      @cell.change_border_top('thin')
      @cell.border_top.should == 'thin'
    end

    it 'should return nil if no top border has been specified for this cell' do
      @cell.border_top.should be_nil
    end
  end

  describe '.border_left' do
    it 'should correctly return the weight of the border on left for this cell' do
      @cell.change_border_left('thin')
      @cell.border_left.should == 'thin'
    end

    it 'should return nil if no left border has been specified for this cell' do
      @cell.border_left.should be_nil
    end
  end

  describe '.border_right' do
    it 'should correctly return the weight of the border on right for this cell' do
      @cell.change_border_right('thin')
      @cell.border_right.should == 'thin'
    end

    it 'should return nil if no right border has been specified for this cell' do
      @cell.border_right.should be_nil
    end
  end

  describe '.border_bottom' do
    it 'should correctly return the weight of the border on bottom for this cell' do
      @cell.change_border_bottom('thin')
      @cell.border_bottom.should == 'thin'
    end

    it 'should return nil if no bottom border has been specified for this cell' do
      @cell.border_bottom.should be_nil
    end
  end

  describe '.border_diagonal' do
    it 'should correctly return the weight of the diagonal border for this cell' do
      @cell.change_border_diagonal('thin')
      @cell.border_diagonal.should == 'thin'
    end

    it 'should return nil if no diagonal border has been specified for this cell' do
      @cell.border_diagonal.should be_nil
    end
  end

  describe '.convert_to_cell' do
    it 'should correctly return the "Excel Style" description of cells when given a row/column number' do
      RubyXL::Cell.convert_to_cell(0,26).should == 'AA1'
    end

    it 'should cause an error if a negative argument is given' do
      lambda {RubyXL::Cell.convert_to_cell(-1,0)}.should raise_error
    end
  end
end
