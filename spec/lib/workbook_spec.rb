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

  describe '.application' do
    it 'should contain default application string' do
      expect(@workbook.application).to eq(RubyXL::Workbook::APPLICATION)
    end

    it 'should set application properly' do
      @workbook.application = 'TEST APPLICATION'
      expect(@workbook.application).to eq('TEST APPLICATION')
    end
  end

  describe '.company' do
    it 'should have default company empty' do
      expect(@workbook.company).to eq('')
    end

    it 'should set company properly' do
      @workbook.company = 'TEST COMPANY'
      expect(@workbook.company).to eq('TEST COMPANY')
    end
  end

  describe '.appversion' do
    it 'should contain default appversion' do
      expect(@workbook.appversion).to eq(RubyXL::Workbook::APPVERSION)
    end

    it 'should set appversion properly' do
      @workbook.appversion = '12.34'
      expect(@workbook.appversion).to eq('12.34')
    end
  end

  describe '.creator' do
    it 'should contain default creator' do
      expect(@workbook.creator).to be_nil
    end

    it 'should set creator properly' do
      @workbook.creator = 'CREATOR'
      expect(@workbook.creator).to eq('CREATOR')
    end
  end

  describe '.modifier' do
    it 'should contain default modifier' do
      expect(@workbook.modifier).to be_nil
    end

    it 'should set modifier properly' do
      @workbook.modifier = 'MODIFIER'
      expect(@workbook.modifier).to eq('MODIFIER')
    end
  end

  describe '.created_at' do
    it 'should contain current time by default' do
      expect(@workbook.created_at).to be_a_kind_of(Time)
    end

    it 'should set modifier properly' do
      dt = Time.at(Time.now.to_i) # Strip time of microseconds
      @workbook.created_at = dt
      expect(@workbook.created_at.to_time).to eq(dt)
    end
  end

  describe '.created_at' do
    it 'should contain current time by default' do
      expect(@workbook.modified_at).to be_a_kind_of(Time)
    end

    it 'should set modifier properly' do
      dt = Time.at(Time.now.to_i) # Strip time of microseconds
      @workbook.modified_at = dt
      expect(@workbook.modified_at.to_time).to eq(dt)
    end
  end

end
