require 'rubygems'
require 'rubyXL'

describe RubyXL::Workbook do
  before do
    @workbook  = RubyXL::Workbook.new
    @worksheet = @workbook.add_worksheet('Test Worksheet')

    (0..10).each do |i|
      (0..10).each do |j|
        @worksheet.add_cell(i, j, "#{i}:#{j}")
      end
    end

    @cell = @worksheet[0][0]
  end

  describe '.new' do
    it 'should automatically create a blank worksheet named "Sheet1"' do
      @workbook[0].sheet_name.should == 'Sheet1'
    end
  end

  describe '[]' do
    it 'should properly locate worksheet by index' do
      @workbook[1].sheet_name.should == 'Test Worksheet'
    end

    it 'should properly locate worksheet by name' do
      @workbook['Test Worksheet'].sheet_name.should == 'Test Worksheet'
    end
  end

  describe '.add_worksheet' do
    it 'when not given a name, it should automatically pick a name "SheetX" that is not taken yet' do
      @workbook.add_worksheet
      @workbook['Sheet2'].sheet_name.should == 'Sheet2'
    end
  end

  describe '.write' do
    #method not conducive to unit tests
  end

  describe '.get_style' do
    it 'should return the cell_xfs object based on the passed in style index (string or number)' do
      @workbook.get_style('0').should == @workbook.cell_xfs[:xf][0]
    end

    it 'should return nil if index out of range or string is passed in' do
      @workbook.get_style('20000').should be_nil
    end
  end

  describe '.get_style_attributes' do
    it 'should return the attributes of the style object when passed the style object itself' do
      @workbook.get_style_attributes(@workbook.get_style(0)).should == @workbook.cell_xfs[:xf][0][:attributes]
    end

    it 'should cause an error if nil is passed' do
      lambda {@workbook.get_style_attributes(nil)}.should raise_error
    end
  end

  describe '.get_fill_color' do
    it 'should return the fill color of a particular style attribute' do
      @cell.change_fill('000000')
      @workbook.get_fill_color(@workbook.get_style_attributes(@workbook.get_style(@cell.style_index))).should == '000000'
    end

    it 'should return white (ffffff) if no fill color is specified in style' do
      @workbook.get_fill_color(@workbook.get_style_attributes(@workbook.get_style(@cell.style_index))).should == 'ffffff'
    end
  end
end
