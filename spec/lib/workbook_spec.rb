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
      expect(@workbook[0]).not_to be_nil
      expect(@workbook[0].sheet_name).to eq('Sheet1')
    end
  end

  describe '[]' do
    it 'should properly locate worksheet by index' do
      expect(@workbook[1]).not_to be_nil
      expect(@workbook[1].sheet_name).to eq('Test Worksheet')
    end

    it 'should properly locate worksheet by name' do
      expect(@workbook['Test Worksheet']).not_to be_nil
      expect(@workbook['Test Worksheet'].sheet_name).to eq('Test Worksheet')
    end
  end

  describe '.add_worksheet' do
    it 'when not given a name, it should automatically pick a name "SheetX" that is not taken yet' do
      expect(@workbook['Sheet2']).to be_nil
      @workbook.add_worksheet
      expect(@workbook['Sheet2']).not_to be_nil
      expect(@workbook['Sheet2'].sheet_name).to eq('Sheet2')
    end
  end

  describe '.get_fill_color' do
    it 'should return the fill color of a particular style attribute' do
      @cell.change_fill('000000')
      expect(@workbook.get_fill_color(@workbook.cell_xfs[@cell.style_index])).to eq('000000')
    end

    it 'should return white (ffffff) if no fill color is specified in style' do
      expect(@workbook.get_fill_color(@workbook.cell_xfs[@cell.style_index])).to eq('ffffff')
    end
  end

end
