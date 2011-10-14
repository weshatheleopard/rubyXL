require 'rubygems'
require 'rubyXL'

describe RubyXL::Worksheet do
  before do
    @workbook  = RubyXL::Workbook.new
    @worksheet = RubyXL::Worksheet.new(@workbook)
    @workbook.worksheets << @worksheet
    (0..10).each do |i|
      (0..10).each do |j|
        @worksheet.add_cell(i, j, "#{i}:#{j}")
      end
    end

    @old_cell = Marshal.load(Marshal.dump(@worksheet[0][0]))
    @old_cell_value = @worksheet[0][0].value
    @old_cell_formula = @worksheet[0][0].formula
  end

  describe '.extract_data' do
    it 'should return a 2d array of just the cell values (without style or formula information)' do
      data = @worksheet.extract_data()
      data[0][0].should == '0:0'
      data.size.should == @worksheet.sheet_data.size
      data[0].size.should == @worksheet[0].size
    end
  end

  describe '.get_table' do
    it 'should return nil if table cannot be found with specified string' do
      @worksheet.get_table('TEST').should be_nil
    end

    it 'should return nil if table cannot be found with specified headers' do
      @worksheet.get_table(['TEST']).should be_nil
    end

    it 'should return a hash when given an array of headers it can find, where :table points to an array of hashes (rows), where each symbol is a column' do
      headers = ["0:0", "0:1", "0:4"]
      table_hash = @worksheet.get_table(headers)

      table_hash[:table].size.should == 10
      table_hash["0:0"].size.should == 10
      table_hash["0:1"].size.should == 10
      table_hash["0:4"].size.should == 10
    end
  end

  describe '.change_row_fill' do
  	it 'should raise error if hex color code not passed' do
  	  lambda {
  	    @worksheet.change_row_fill(0, 'G')
  	  }.should raise_error
    end

    it 'should raise error if hex color code includes # character' do
      lambda {
        @worksheet.change_row_fill(3,'#FFF000')
      }.should raise_error
    end

  	it 'should make row and cell fill colors equal hex color code passed' do
  	  @worksheet.change_row_fill(0, '111111')
      @worksheet.get_row_fill(0).should == '111111'
      @worksheet[0][5].fill_color.should == '111111'
  	end

  	it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_fill(-1,'111111')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_fill(11,'111111')
      @worksheet.get_row_fill(11).should == '111111'
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_font_name' do
    it 'should make row and cell font names equal font name passed' do
      @worksheet.change_row_font_name(0, 'Arial')
      @worksheet.get_row_font_name(0).should == 'Arial'
      @worksheet[0][5].font_name.should == 'Arial'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_font_name(-1,'Arial')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_font_name(11, 'Arial')
      @worksheet.get_row_font_name(11).should == "Arial"
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_font_size' do
    it 'should make row and cell font sizes equal font number passed' do
      @worksheet.change_row_font_size(0, 20)
      @worksheet.get_row_font_size(0).should == 20
      @worksheet[0][5].font_size.should == 20
    end

    it 'should cause an error if a string passed' do
      lambda {
        @worksheet.change_row_font_size(0, '20')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_font_size(-1,20)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_font_size(11,20)
      @worksheet.get_row_font_size(11).should == 20
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_font_color' do
    it 'should make row and cell font colors equal to font color passed' do
      @worksheet.change_row_font_color(0, '0f0f0f')
      @worksheet.get_row_font_color(0).should == '0f0f0f'
      @worksheet[0][5].font_color.should == '0f0f0f'
    end

    it 'should raise error if hex color code not passed' do
  	  lambda {
  	    @worksheet.change_row_font_color(0, 'G')
  	  }.should raise_error
    end

    it 'should raise error if hex color code includes # character' do
      lambda {
        @worksheet.change_row_font_color(3,'#FFF000')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_font_color(-1,'0f0f0f')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_font_color(11,'0f0f0f')
      @worksheet.get_row_font_color(11).should == '0f0f0f'
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_italics' do
    it 'should make row and cell fonts italicized when true is passed' do
      @worksheet.change_row_italics(0,true)
      @worksheet.is_row_italicized(0).should == true
      @worksheet[0][5].is_italicized.should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_italics(-1,false)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_italics(11,true)
      @worksheet.is_row_italicized(11).should == true
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_bold' do
    it 'should make row and cell fonts bolded when true is passed' do
      @worksheet.change_row_bold(0,true)
      @worksheet.is_row_bolded(0).should == true
      @worksheet[0][5].is_bolded.should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_bold(-1,false)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_bold(11,true)
      @worksheet.is_row_bolded(11).should == true
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_underline' do
    it 'should make row and cell fonts underlined when true is passed' do
      @worksheet.change_row_underline(0,true)
      @worksheet.is_row_underlined(0).should == true
      @worksheet[0][5].is_underlined.should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_underline(-1,false)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_underline(11,true)
      @worksheet.is_row_underlined(11).should == true
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_strikethrough' do
    it 'should make row and cell fonts struckthrough when true is passed' do
      @worksheet.change_row_strikethrough(0,true)
      @worksheet.is_row_struckthrough(0).should == true
      @worksheet[0][5].is_struckthrough.should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_strikethrough(-1,false)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_strikethrough(11,true)
      @worksheet.is_row_struckthrough(11).should == true
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_height' do
    it 'should make row height match number which is passed' do
      @worksheet.change_row_height(0,30.0002)
      @worksheet.get_row_height(0).should == 30.0002
    end

    it 'should make row height a number equivalent of the string passed if it is a string which is a number' do
      @worksheet.change_row_height(0,'30.0002')
      @worksheet.get_row_height(0).should == 30.0002
    end

    it 'should cause error if a string which is not a number' do
      lambda {
        @worksheet.change_row_height(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_height(-1,30)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_height(11,30)
      @worksheet.get_row_height(11).should == 30
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_horizontal_alignment' do
    it 'should cause row and cells to horizontally align as specified by the passed in string' do
      @worksheet.change_row_horizontal_alignment(0,'center')
      @worksheet.get_row_horizontal_alignment(0).should == 'center'
      @worksheet[0][5].horizontal_alignment.should == 'center'
    end

    it 'should cause error if nil, "center", "justify", "left", "right", or "distributed" is not passed' do
      lambda {
        @worksheet.change_row_horizontal_alignment(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_horizontal_alignment(-1,'center')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_horizontal_alignment(11,'center')
      @worksheet.get_row_horizontal_alignment(11).should == 'center'
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_vertical_alignment' do
    it 'should cause row and cells to vertically align as specified by the passed in string' do
      @worksheet.change_row_vertical_alignment(0,'center')
      @worksheet.get_row_vertical_alignment(0).should == 'center'
      @worksheet[0][5].vertical_alignment.should == 'center'
    end

    it 'should cause error if nil, "center", "justify", "top", "bottom", or "distributed" is not passed' do
      lambda {
        @worksheet.change_row_vertical_alignment(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_vertical_alignment(-1,'center')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_vertical_alignment(11,'center')
      @worksheet.get_row_vertical_alignment(11).should == 'center'
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_border_top' do
    it 'should cause row and cells to have border at top of specified weight' do
      @worksheet.change_row_border_top(0, 'thin')
      @worksheet.get_row_border_top(0).should == 'thin'
      @worksheet[0][5].border_top.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @worksheet.change_row_border_top(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_border_top(-1,'thin')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_border_top(11,'thin')
      @worksheet.get_row_border_top(11).should == 'thin'
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_border_left' do
    it 'should cause row and cells to have border at left of specified weight' do
      @worksheet.change_row_border_left(0, 'thin')
      @worksheet.get_row_border_left(0).should == 'thin'
      @worksheet[0][5].border_left.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @worksheet.change_row_border_left(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_border_left(-1,'thin')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'  do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_border_left(11,'thin')
      @worksheet.get_row_border_left(11).should == 'thin'
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_border_right' do
    it 'should cause row and cells to have border at right of specified weight' do
      @worksheet.change_row_border_right(0, 'thin')
      @worksheet.get_row_border_right(0).should == 'thin'
      @worksheet[0][5].border_right.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @worksheet.change_row_border_right(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_border_right(-1,'thin')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_border_right(11,'thin')
      @worksheet.get_row_border_right(11).should == 'thin'
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_border_bottom' do
    it 'should cause row to have border at bottom of specified weight' do
      @worksheet.change_row_border_bottom(0, 'thin')
      @worksheet.get_row_border_bottom(0).should == 'thin'
      @worksheet[0][5].border_bottom.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @worksheet.change_row_border_bottom(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_border_bottom(-1,'thin')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_border_bottom(11,'thin')
      @worksheet.get_row_border_bottom(11).should == 'thin'
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_row_border_diagonal' do
    it 'should cause row to have border at diagonal of specified weight' do
      @worksheet.change_row_border_diagonal(0, 'thin')
      @worksheet.get_row_border_diagonal(0).should == 'thin'
      @worksheet[0][5].border_diagonal.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @worksheet.change_row_border_diagonal(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_row_border_diagonal(-1,'thin')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      @worksheet.change_row_border_diagonal(11,'thin')
      @worksheet.get_row_border_diagonal(11).should == 'thin'
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.change_column_font_name' do
    it 'should cause column and cell font names to match string passed in' do
      @worksheet.change_column_font_name(0, 'Arial')
      @worksheet.get_column_font_name(0).should == 'Arial'
      @worksheet[5][0].font_name.should == 'Arial'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_font_name(-1,'Arial')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_font_name(11,'Arial')
      @worksheet.get_column_font_name(11).should == 'Arial'
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_font_size' do
    it 'should make column and cell font sizes equal font number passed' do
      @worksheet.change_column_font_size(0, 20)
      @worksheet.get_column_font_size(0).should == 20
      @worksheet[5][0].font_size.should == 20
    end

    it 'should cause an error if a string passed' do
      lambda {
        @worksheet.change_column_font_size(0, '20')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_font_size(-1,20)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_font_size(11,20)
      @worksheet.get_column_font_size(11).should == 20
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_font_color' do
    it 'should make column and cell font colors equal to font color passed' do
      @worksheet.change_column_font_color(0, '0f0f0f')
      @worksheet.get_column_font_color(0).should == '0f0f0f'
      @worksheet[5][0].font_color.should == '0f0f0f'
    end

    it 'should raise error if hex color code not passed' do
  	  lambda {
  	    @worksheet.change_column_font_color(0, 'G')
  	  }.should raise_error
    end

    it 'should raise error if hex color code includes # character' do
      lambda {
        @worksheet.change_column_font_color(0,'#FFF000')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_font_color(-1,'0f0f0f')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_font_color(11,'0f0f0f')
      @worksheet.get_column_font_color(11).should == '0f0f0f'
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_italics' do
    it 'should make column and cell fonts italicized when true is passed' do
      @worksheet.change_column_italics(0,true)
      @worksheet.is_column_italicized(0).should == true
      @worksheet[5][0].is_italicized.should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_italicized(-1,false)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_italics(11,true)
      @worksheet.is_column_italicized(11).should == true
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_bold' do
    it 'should make column and cell fonts bolded when true is passed' do
      @worksheet.change_column_bold(0,true)
      @worksheet.is_column_bolded(0).should == true
      @worksheet[5][0].is_bolded.should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_bold(-1,false)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_bold(11,true)
      @worksheet.is_column_bolded(11).should == true
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_underline' do
    it 'should make column and cell fonts underlined when true is passed' do
      @worksheet.change_column_underline(0,true)
      @worksheet.is_column_underlined(0).should == true
      @worksheet[5][0].is_underlined.should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_underline(-1,false)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_underline(11,true)
      @worksheet.is_column_underlined(11).should == true
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_strikethrough' do
    it 'should make column and cell fonts struckthrough when true is passed' do
      @worksheet.change_column_strikethrough(0,true)
      @worksheet.is_column_struckthrough(0).should == true
      @worksheet[5][0].is_struckthrough.should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_strikethrough(-1,false)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_strikethrough(11,true)
      @worksheet.is_column_struckthrough(11).should == true
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_width' do
    it 'should make column width match number which is passed' do
      @worksheet.change_column_width(0,30.0002)
      @worksheet.get_column_width(0).should == 30.0002
    end

    it 'should make column width a number equivalent of the string passed if it is a string which is a number' do
      @worksheet.change_column_width(0,'30.0002')
      @worksheet.get_column_width(0).should == 30.0002
    end

    it 'should cause error if a string which is not a number' do
      lambda {
        @worksheet.change_column_width(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_width(-1,10)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_width(11,30)
      @worksheet.get_column_width(11).should == 30
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_fill' do
    it 'should raise error if hex color code not passed' do
  	  lambda {
  	    @worksheet.change_column_fill(0, 'G')
  	  }.should raise_error
    end

    it 'should raise error if hex color code includes # character' do
      lambda {
        @worksheet.change_column_fill(3,'#FFF000')
      }.should raise_error
    end

  	it 'should make column and cell fill colors equal hex color code passed' do
  	  @worksheet.change_column_fill(0, '111111')
      @worksheet.get_column_fill(0).should == '111111'
      @worksheet[5][0].fill_color.should == '111111'
  	end

  	it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_fill(-1,'111111')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_fill(11,'111111')
      @worksheet.get_column_fill(11).should == '111111'
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_horizontal_alignment' do
    it 'should cause column and cell to horizontally align as specified by the passed in string' do
      @worksheet.change_column_horizontal_alignment(0,'center')
      @worksheet.get_column_horizontal_alignment(0).should == 'center'
      @worksheet[5][0].horizontal_alignment.should == 'center'
    end

    it 'should cause error if nil, "center", "justify", "left", "right", or "distributed" is not passed' do
      lambda {
        @worksheet.change_column_horizontal_alignment(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_horizontal_alignment(-1,'center')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_horizontal_alignment(11,'center')
      @worksheet.get_column_horizontal_alignment(11).should == 'center'
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_vertical_alignment' do
    it 'should cause column and cell to vertically align as specified by the passed in string' do
      @worksheet.change_column_vertical_alignment(0,'center')
      @worksheet.get_column_vertical_alignment(0).should == 'center'
      @worksheet[5][0].vertical_alignment.should == 'center'
    end

    it 'should cause error if nil, "center", "justify", "top", "bottom", or "distributed" is not passed' do
      lambda {
        @worksheet.change_column_vertical_alignment(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_vertical_alignment(-1,'center')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_vertical_alignment(11,'center')
      @worksheet.get_column_vertical_alignment(11).should == 'center'
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_border_top' do
    it 'should cause column and cells within to have border at top of specified weight' do
      @worksheet.change_column_border_top(0, 'thin')
      @worksheet.get_column_border_top(0).should == 'thin'
      @worksheet[5][0].border_top.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @worksheet.change_column_border_top(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_border_top(-1,'thin')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_border_top(11,'thin')
      @worksheet.get_column_border_top(11).should == 'thin'
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_border_left' do
    it 'should cause column and cells within to have border at left of specified weight' do
      @worksheet.change_column_border_left(0, 'thin')
      @worksheet.get_column_border_left(0).should == 'thin'
      @worksheet[5][0].border_left.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @worksheet.change_column_border_left(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_border_left(-1,'thin')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_border_left(11,'thin')
      @worksheet.get_column_border_left(11).should == 'thin'
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_border_right' do
    it 'should cause column and cells within to have border at right of specified weight' do
      @worksheet.change_column_border_right(0, 'thin')
      @worksheet.get_column_border_right(0).should == 'thin'
      @worksheet[5][0].border_right.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @worksheet.change_column_border_right(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_border_right(-1,'thin')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_border_right(11,'thin')
      @worksheet.get_column_border_right(11).should == 'thin'
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_border_bottom' do
    it 'should cause column and cells within to have border at bottom of specified weight' do
      @worksheet.change_column_border_bottom(0, 'thin')
      @worksheet.get_column_border_bottom(0).should == 'thin'
      @worksheet[5][0].border_bottom.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @worksheet.change_column_border_bottom(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_border_bottom(-1,'thin')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_border_bottom(11,'thin')
      @worksheet.get_column_border_bottom(11).should == 'thin'
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.change_column_border_diagonal' do
    it 'should cause column and cells within to have border at diagonal of specified weight' do
      @worksheet.change_column_border_diagonal(0, 'thin')
      @worksheet.get_column_border_diagonal(0).should == 'thin'
      @worksheet[5][0].border_diagonal.should == 'thin'
    end

    it 'should cause error if nil, "thin", "thick", "hairline", or "medium" is not passed' do
      lambda {
        @worksheet.change_column_border_diagonal(0,'TEST')
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.change_column_border_diagonal(-1,'thin')
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.change_column_border_diagonal(11,'thin')
      @worksheet.get_column_border_diagonal(11).should == 'thin'
      @worksheet.sheet_data[0].size.should == 12
    end
  end

  describe '.merge_cells' do
    it 'should merge cells in any valid range specified by indices' do
      @worksheet.merge_cells(0,0,1,1)
      @worksheet.merged_cells.include?({:attributes=>{:ref=>"A1:B2"}}).should == true
    end

    it 'should cause an error if a negative number is passed' do
      lambda {
        @worksheet.merge_cells(0,0,-1,0)
      }.should raise_error
    end
  end

  describe '.add_cell' do
    it 'should add new cell where specified, even if a cell is already there (default)' do
      @worksheet.add_cell(0,0,'TEST')
      @worksheet[0][0].value.should_not == @old_cell_value
      @worksheet[0][0].value.should == 'TEST'
    end

    it 'should add new cell where specified with formula, even if a cell is already there (default)' do
      @worksheet.add_cell(0,0,'','SUM(A2:A10)')
      @worksheet[0][0].value.should_not == @old_cell_value
      @worksheet[0][0].formula.should_not == @old_cell_formula
      @worksheet[0][0].value.should == ''
      @worksheet[0][0].formula.should == 'SUM(A2:A10)'
    end

    it 'should not overwrite when a cell is present when overwrite is specified to be false' do
      @worksheet.add_cell(0,0,'TEST','B2',false)
      @worksheet[0][0].value.should == @old_cell_value
      @worksheet[0][0].formula.should == @old_cell_formula
    end

    it 'should still add a new cell when there is no cell to be overwritten' do
      @worksheet.add_cell(11,11,'TEST','B2',false)
      @worksheet[11][11].value.should == 'TEST'
      @worksheet[11][11].formula.should == 'B2'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.add_cell(-1,-1,'')
      }.should raise_error
    end
  end

  describe '.add_cell_obj' do
    it 'should add already created cell object to worksheet, even if a cell is already there (default)' do
      new_cell = RubyXL::Cell.new(@worksheet,0,0,'TEST','B2')
      @worksheet.add_cell_obj(new_cell)
      @worksheet[0][0].value.should_not == @old_cell_value
      @worksheet[0][0].formula.should_not == @old_cell_formula
      @worksheet[0][0].value.should == 'TEST'
      @worksheet[0][0].formula.should == 'B2'
    end

    it 'should not add already created cell object to already occupied cell if overwrite is false' do
      new_cell = RubyXL::Cell.new(@worksheet,0,0,'TEST','B2')
      @worksheet.add_cell_obj(new_cell,false)
      @worksheet[0][0].value.should == @old_cell_value
      @worksheet[0][0].formula.should == @old_cell_formula
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.add_cell_obj(-1)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data.size.should == 11
      new_cell = RubyXL::Cell.new(@worksheet,11,11,'TEST','B2')
      @worksheet.add_cell_obj(new_cell)
      @worksheet.sheet_data.size.should == 12
    end
  end

  describe '.delete_row' do
    it 'should delete a row at index specified, "pushing" everything else "up"' do
      @worksheet.delete_row(0)
      @worksheet[0][0].value.should == "1:0"
      @worksheet[0][0].formula.should be_nil
      @worksheet[0][0].row.should == 0
      @worksheet[0][0].column.should == 0
    end

    it 'should delete a row at index specified, adjusting styles for other rows' do
      @worksheet.change_row_font_name(1,"Courier")
      @worksheet.delete_row(0)
      @worksheet.get_row_font_name(0).should == "Courier"
    end

    it 'should preserve (rather than fix) formulas that reference cells in "pushed up" rows' do
      @worksheet.add_cell(11,0,nil,'SUM(A1:A10)')
      @worksheet.delete_row(0)
      @worksheet[10][0].formula.should == 'SUM(A1:A10)'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.delete_row(-1)
      }.should raise_error
    end
  end

  describe '.insert_row' do
    it 'should insert a row at index specified, "pushing" everything else "down"' do
      @worksheet.insert_row(0)
      @worksheet[0][0].should be_nil
      @worksheet[1][0].value.should == @old_cell_value
      @worksheet[1][0].formula.should == @old_cell_formula
    end

    it 'should insert a row at index specified, copying styles from row "above"' do
      @worksheet.change_row_font_name(0,'Courier')
      @worksheet.insert_row(1)
      @worksheet.get_row_font_name(1).should == 'Courier'
    end

    it 'should preserve (rather than fix) formulas that reference cells "pushed down" rows' do
      @worksheet.add_cell(5,0,nil,'SUM(A1:A4)')
      @worksheet.insert_row(0)
      @worksheet[6][0].formula.should == 'SUM(A1:A4)'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.insert_row(-1)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative'  do
      @worksheet.sheet_data.size.should == 11
      @worksheet.insert_row(11)
      @worksheet.sheet_data.size.should == 13
    end
  end

  describe '.delete_column' do
    it 'should delete a column at index specified, "pushing" everything else "left"' do
      @worksheet.delete_column(0)
      @worksheet[0][0].value.should == "0:1"
      @worksheet[0][0].formula.should be_nil
      @worksheet[0][0].row.should == 0
      @worksheet[0][0].column.should == 0
    end

    it 'should delete a column at index specified, "pushing" styles "left"' do
      @worksheet.change_column_font_name(1,"Courier")
      @worksheet.delete_column(0)
      @worksheet.get_column_font_name(0).should == "Courier"
    end

    it 'should preserve (rather than fix) formulas that reference cells in "pushed left" columns' do
      @worksheet.add_cell(0,4,nil,'SUM(A1:D1)')
      @worksheet.delete_column(0)
      @worksheet[0][3].formula.should == 'SUM(A1:D1)'
    end

    it 'should cause error if negative argument is passed in' do
      lambda {
        @worksheet.delete_column(-1)
      }.should raise_error
    end
  end

  describe '.insert_column' do
    it 'should insert a column at index specified, "pushing" everything else "right"' do
      @worksheet.insert_column(0)
      @worksheet[0][0].should be_nil
      @worksheet[0][1].value.should == @old_cell_value
      @worksheet[0][1].formula.should == @old_cell_formula
    end

    it 'should insert a column at index specified, copying styles from column to "left"' do
      @worksheet.change_column_font_name(0,'Courier')
      @worksheet.insert_column(1)
      @worksheet.get_column_font_name(1).should == 'Courier'
    end

    it 'should insert a column at 0 without copying any styles, when passed 0 as column index' do
      @worksheet.change_column_font_name(0,'Courier')
      @worksheet.insert_column(0)
      @worksheet.get_column_font_name(0).should == 'Verdana' #not courier
    end

    it 'should preserve (rather than fix) formulas that reference cells in "pushed right" column' do
      @worksheet.add_cell(0,5,nil,'SUM(A1:D1)')
      @worksheet.insert_column(0)
      @worksheet[0][6].formula.should == 'SUM(A1:D1)'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.insert_column(-1)
      }.should raise_error
    end

    it 'should expand matrix to fit argument if nonnegative' do
      @worksheet.sheet_data[0].size.should == 11
      @worksheet.insert_column(11)
      @worksheet.sheet_data[0].size.should == 13
    end
  end

  describe '.insert_cell' do
    it 'should simply add a cell if no shift argument is specified' do
      @worksheet.insert_cell(0,0,'test')
      @worksheet[0][0].value.should == 'test'
      @worksheet[0][1].value.should == '0:1'
      @worksheet[1][0].value.should == '1:0'
    end

    it 'should shift cells to the right if :right is specified' do
      @worksheet.insert_cell(0,0,'test',nil,:right)
      @worksheet[0][0].value.should == 'test'
      @worksheet[0][1].value.should == '0:0'
      @worksheet[1][0].value.should == '1:0'
    end

    it 'should shift cells down if :down is specified' do
      @worksheet.insert_cell(0,0,'test',nil,:down)
      @worksheet[0][0].value.should == 'test'
      @worksheet[0][1].value.should == '0:1'
      @worksheet[1][0].value.should == '0:0'
    end

    it 'should cause error if shift argument is specified whcih is not :right or :down' do
      lambda {
        @worksheet.insert_cell(0,0,'test',nil,:up)
      }.should raise_error
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.insert_cell(-1,-1)
      }.should raise_error
    end
  end

  describe '.delete_cell' do
    it 'should make a cell nil if no shift argument specified' do
      deleted = @worksheet.delete_cell(0,0)
      @worksheet[0][0].should be_nil
      @old_cell.inspect.should == deleted.inspect
    end

    it 'should return nil if a cell which is out of range is specified' do
      @worksheet.delete_cell(12,12).should be_nil
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.delete_cell(-1,-1)
      }.should raise_error
    end

    it 'should shift cells to the right of the deleted cell left if :left is specified' do
      @worksheet.delete_cell(0,0,:left)
      @worksheet[0][0].value.should == '0:1'
    end

    it 'should shift cells below the deleted cell up if :up is specified' do
      @worksheet.delete_cell(0,0,:up)
      @worksheet[0][0].value.should == '1:0'
    end

    it 'should cause en error if an argument other than :left, :up, or nil is specified for shift' do
      lambda {
        @worksheet.delete_cell(0,0,:down)
      }.should raise_error
    end
  end

  describe '.get_row_fill' do
    it 'should return white (ffffff) if no fill color specified for row' do
      @worksheet.get_row_fill(0).should == 'ffffff'
    end

    it 'should correctly reflect fill color if specified for row' do
      @worksheet.change_row_fill(0, '000000')
      @worksheet.get_row_fill(0).should == '000000'
    end

    it 'should return nil if a row which does not exist is passed in' do
      @worksheet.get_row_fill(11).should be_nil
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_fill(-1)
      }.should raise_error
    end
  end

  describe '.get_row_font_name' do
    it 'should correctly reflect font name for row' do
      @worksheet.change_row_font_name(0,'Courier')
      @worksheet.get_row_font_name(0).should == 'Courier'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_font_name(-1)
      }.should raise_error
    end

    it 'should return nil if a (nonnegative) row which does not exist is passed in' do
      @worksheet.get_row_font_name(11).should be_nil
    end
  end

  describe '.get_row_font_size' do
    it 'should correctly reflect font size for row' do
      @worksheet.change_row_font_size(0,30)
      @worksheet.get_row_font_size(0).should == 30
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_font_size(-1)
      }.should raise_error
    end

    it 'should return nil if a (nonnegative) row which does not exist is passed in' do
      @worksheet.get_row_font_size(11).should be_nil
    end
  end

  describe '.get_row_font_color' do
    it 'should correctly reflect font color for row' do
      @worksheet.change_row_font_color(0,'0f0f0f')
      @worksheet.get_row_font_color(0).should == '0f0f0f'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_font_color(-1)
      }.should raise_error
    end

    it 'should return nil if a (nonnegative) row which does not exist is passed in' do
      @worksheet.get_row_font_color(11).should be_nil
    end
  end

  describe '.is_row_italicized' do
    it 'should correctly return whether row is italicized' do
      @worksheet.change_row_italics(0,true)
      @worksheet.is_row_italicized(0).should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.is_row_italicized(-1)
      }.should raise_error
    end

    it 'should return nil if a (nonnegative) row which does not exist is passed in' do
      @worksheet.is_row_italicized(11).should be_nil
    end
  end

  describe '.is_row_bolded' do
    it 'should correctly return whether row is bolded' do
      @worksheet.change_row_bold(0,true)
      @worksheet.is_row_bolded(0).should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.is_row_bolded(-1)
      }.should raise_error
    end

    it 'should return nil if a (nonnegative) row which does not exist is passed in' do
      @worksheet.is_row_bolded(11).should be_nil
    end
  end

  describe '.is_row_underlined' do
    it 'should correctly return whether row is underlined' do
      @worksheet.change_row_underline(0,true)
      @worksheet.is_row_underlined(0).should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.is_row_underlined(-1)
      }.should raise_error
    end

    it 'should return nil if a (nonnegative) row which does not exist is passed in' do
      @worksheet.is_row_underlined(11).should be_nil
    end
  end

  describe '.is_row_struckthrough' do
    it 'should correctly return whether row is struckthrough' do
      @worksheet.change_row_strikethrough(0,true)
      @worksheet.is_row_struckthrough(0).should == true
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.is_row_struckthrough(-1)
      }.should raise_error
    end

    it 'should return nil if a (nonnegative) row which does not exist is passed in' do
      @worksheet.is_row_struckthrough(11).should be_nil
    end
  end

  describe '.get_row_height' do
    it 'should return 13 if no height specified for row' do
      @worksheet.get_row_height(0).should == 13
    end

    it 'should correctly reflect height if specified for row' do
      @worksheet.change_row_height(0, 30)
      @worksheet.get_row_height(0).should == 30
    end

    it 'should return nil if a row which does not exist is passed in' do
      @worksheet.get_row_height(11).should be_nil
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_height(-1)
      }.should raise_error
    end
  end

  describe '.get_row_horizontal_alignment' do
    it 'should return nil if no alignment specified for row' do
      @worksheet.get_row_horizontal_alignment(0).should be_nil
    end

    it 'should return nil if a row which does not exist is passed in' do
      @worksheet.get_row_horizontal_alignment(11).should be_nil
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_horizontal_alignment(-1)
      }.should raise_error
    end

    it 'should return correct horizontal alignment if it is set for that row' do
      @worksheet.change_row_horizontal_alignment(0, 'center')
      @worksheet.get_row_horizontal_alignment(0).should == 'center'
    end
  end

  describe '.get_row_vertical_alignment' do
    it 'should return nil if no alignment specified for row' do
      @worksheet.get_row_vertical_alignment(0).should be_nil
    end

    it 'should return nil if a row which does not exist is passed in' do
      @worksheet.get_row_vertical_alignment(11).should be_nil
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_vertical_alignment(-1)
      }.should raise_error
    end

    it 'should return correct vertical alignment if it is set for that row' do
      @worksheet.change_row_vertical_alignment(0, 'center')
      @worksheet.get_row_vertical_alignment(0).should == 'center'
    end
  end

  describe '.get_row_border_top' do
    it 'should return nil if no border is specified for that row in that direction' do
      @worksheet.get_row_border_top(0).should be_nil
    end

    it 'should return type of border that this row has on top' do
      @worksheet.change_row_border_top(0,'thin')
      @worksheet.get_row_border_top(0).should == 'thin'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_border_top(-1)
      }.should raise_error
    end

    it 'should return nil if a row which does not exist is passed in' do
      @worksheet.get_row_border_top(11).should be_nil
    end
  end

  describe '.get_row_border_left' do
    it 'should return nil if no border is specified for that row in that direction' do
      @worksheet.get_row_border_left(0).should be_nil
    end

    it 'should return type of border that this row has on left' do
      @worksheet.change_row_border_left(0,'thin')
      @worksheet.get_row_border_left(0).should == 'thin'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_border_left(-1)
      }.should raise_error
    end

    it 'should return nil if a row which does not exist is passed in' do
      @worksheet.get_row_border_left(11).should be_nil
    end
  end

  describe '.get_row_border_right' do
    it 'should return nil if no border is specified for that row in that direction' do
      @worksheet.get_row_border_right(0).should be_nil
    end

    it 'should return type of border that this row has on right' do
      @worksheet.change_row_border_right(0,'thin')
      @worksheet.get_row_border_right(0).should == 'thin'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_border_right(-1)
      }.should raise_error
    end

    it 'should return nil if a row which does not exist is passed in' do
      @worksheet.get_row_border_right(11).should be_nil
    end
  end


  describe '.get_row_border_bottom' do
    it 'should return nil if no border is specified for that row in that direction' do
      @worksheet.get_row_border_bottom(0).should be_nil
    end

    it 'should return type of border that this row has on bottom' do
      @worksheet.change_row_border_bottom(0,'thin')
      @worksheet.get_row_border_bottom(0).should == 'thin'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_border_bottom(-1)
      }.should raise_error
    end

    it 'should return nil if a row which does not exist is passed in' do
      @worksheet.get_row_border_bottom(11).should be_nil
    end
  end

  describe '.get_row_border_diagonal' do
    it 'should return nil if no border is specified for that row in that direction' do
      @worksheet.get_row_border_diagonal(0).should be_nil
    end

    it 'should return type of border that this row has on diagonal' do
      @worksheet.change_row_border_diagonal(0,'thin')
      @worksheet.get_row_border_diagonal(0).should == 'thin'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_row_border_diagonal(-1)
      }.should raise_error
    end

    it 'should return nil if a row which does not exist is passed in' do
      @worksheet.get_row_border_diagonal(11).should be_nil
    end
  end

  describe '.get_column_font_name' do
    it 'should correctly reflect font name for column' do
      @worksheet.change_column_font_name(0,'Courier')
      @worksheet.get_column_font_name(0).should == 'Courier'
    end

    it 'should cause error if a negative argument is passed in' do
      lambda {
        @worksheet.get_column_font_name(-1)
      }.should raise_error
    end

    it 'should return nil if a (nonnegative) column which does not exist is passed in' do
      @worksheet.get_column_font_name(11).should be_nil
    end
  end

  describe '.get_column_font_size' do
     it 'should correctly reflect font size for column' do
       @worksheet.change_column_font_size(0,30)
       @worksheet.get_column_font_size(0).should == 30
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.get_column_font_size(-1)
       }.should raise_error
     end

     it 'should return nil if a (nonnegative) column which does not exist is passed in' do
       @worksheet.get_column_font_size(11).should be_nil
     end
   end

   describe '.get_column_font_color' do
     it 'should correctly reflect font color for column' do
       @worksheet.change_column_font_color(0,'0f0f0f')
       @worksheet.get_column_font_color(0).should == '0f0f0f'
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.get_column_font_color(-1)
       }.should raise_error
     end

     it 'should return nil if a (nonnegative) column which does not exist is passed in' do
       @worksheet.get_column_font_color(11).should be_nil
     end

     it 'should return black (000000) if no rgb font color is specified' do
       @worksheet.get_column_font_color(0).should == '000000'
     end
   end

   describe '.is_column_italicized' do
     it 'should correctly return whether column is italicized' do
       @worksheet.change_column_italics(0,true)
       @worksheet.is_column_italicized(0).should == true
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.is_column_italicized(-1)
       }.should raise_error
     end

     it 'should return nil if a (nonnegative) column which does not exist is passed in' do
       @worksheet.is_column_italicized(11).should be_nil
     end
   end

   describe '.is_column_bolded' do
     it 'should correctly return whether column is bolded' do
       @worksheet.change_column_bold(0,true)
       @worksheet.is_column_bolded(0).should == true
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.is_column_bolded(-1)
       }.should raise_error
     end

     it 'should return nil if a (nonnegative) column which does not exist is passed in' do
       @worksheet.is_column_bolded(11).should be_nil
     end
   end

   describe '.is_column_underlined' do
     it 'should correctly return whether column is underlined' do
       @worksheet.change_column_underline(0,true)
       @worksheet.is_column_underlined(0).should == true
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.is_column_underlined(-1)
       }.should raise_error
     end

     it 'should return nil if a (nonnegative) column which does not exist is passed in' do
       @worksheet.is_column_underlined(11).should be_nil
     end
   end

   describe '.is_column_struckthrough' do
     it 'should correctly return whether column is struckthrough' do
       @worksheet.change_column_strikethrough(0,true)
       @worksheet.is_column_struckthrough(0).should == true
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.is_column_struckthrough(-1)
       }.should raise_error
     end

     it 'should return nil if a (nonnegative) column which does not exist is passed in' do
       @worksheet.is_column_struckthrough(11).should be_nil
     end
   end

   describe '.get_column_width' do
     it 'should return 10 (base column width) if no width specified for column' do
       @worksheet.get_column_width(0).should == 10
     end

     it 'should correctly reflect width if specified for column' do
       @worksheet.change_column_width(0, 30)
       @worksheet.get_column_width(0).should == 30
     end

     it 'should return nil if a column which does not exist is passed in' do
       @worksheet.get_column_width(11).should be_nil
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.get_column_width(-1)
       }.should raise_error
     end
   end

   describe '.get_column_fill' do
     it 'should return white (ffffff) if no fill color specified for column' do
       @worksheet.get_column_fill(0).should == 'ffffff'
     end

     it 'should correctly reflect fill color if specified for column' do
       @worksheet.change_column_fill(0, '000000')
       @worksheet.get_column_fill(0).should == '000000'
     end

     it 'should return nil if a column which does not exist is passed in' do
       @worksheet.get_column_fill(11).should be_nil
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.get_column_fill(-1)
       }.should raise_error
     end
   end

   describe '.get_column_horizontal_alignment' do
     it 'should return nil if no alignment specified for column' do
       @worksheet.get_column_horizontal_alignment(0).should be_nil
     end

     it 'should return nil if a column which does not exist is passed in' do
       @worksheet.get_column_horizontal_alignment(11).should be_nil
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.get_column_horizontal_alignment(-1)
       }.should raise_error
     end

     it 'should return correct horizontal alignment if it is set for that column' do
       @worksheet.change_column_horizontal_alignment(0, 'center')
       @worksheet.get_column_horizontal_alignment(0).should == 'center'
     end
   end

   describe '.get_column_vertical_alignment' do
     it 'should return nil if no alignment specified for column' do
       @worksheet.get_column_vertical_alignment(0).should be_nil
     end

     it 'should return nil if a column which does not exist is passed in' do
       @worksheet.get_column_vertical_alignment(11).should be_nil
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.get_column_vertical_alignment(-1)
       }.should raise_error
     end

     it 'should return correct vertical alignment if it is set for that column' do
       @worksheet.change_column_vertical_alignment(0, 'center')
       @worksheet.get_column_vertical_alignment(0).should == 'center'
     end
   end

   describe '.get_column_border_top' do
     it 'should return nil if no border is specified for that column in that direction' do
       @worksheet.get_column_border_top(0).should be_nil
     end

     it 'should return type of border that this column has on top' do
       @worksheet.change_column_border_top(0,'thin')
       @worksheet.get_column_border_top(0).should == 'thin'
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.get_column_border_top(-1)
       }.should raise_error
     end

     it 'should return nil if a column which does not exist is passed in' do
       @worksheet.get_column_border_top(11).should be_nil
     end
   end

   describe '.get_column_border_left' do
     it 'should return nil if no border is specified for that column in that direction' do
       @worksheet.get_column_border_left(0).should be_nil
     end

     it 'should return type of border that this column has on left' do
       @worksheet.change_column_border_left(0,'thin')
       @worksheet.get_column_border_left(0).should == 'thin'
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.get_column_border_left(-1)
       }.should raise_error
     end

     it 'should return nil if a column which does not exist is passed in' do
       @worksheet.get_column_border_left(11).should be_nil
     end
   end

   describe '.get_column_border_right' do
     it 'should return nil if no border is specified for that column in that direction' do
       @worksheet.get_column_border_right(0).should be_nil
     end

     it 'should return type of border that this column has on right' do
       @worksheet.change_column_border_right(0,'thin')
       @worksheet.get_column_border_right(0).should == 'thin'
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.get_column_border_right(-1)
       }.should raise_error
     end

     it 'should return nil if a column which does not exist is passed in' do
       @worksheet.get_column_border_right(11).should be_nil
     end
   end


   describe '.get_column_border_bottom' do
     it 'should return nil if no border is specified for that column in that direction' do
       @worksheet.get_column_border_bottom(0).should be_nil
     end

     it 'should return type of border that this column has on bottom' do
       @worksheet.change_column_border_bottom(0,'thin')
       @worksheet.get_column_border_bottom(0).should == 'thin'
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.get_column_border_bottom(-1)
       }.should raise_error
     end

     it 'should return nil if a column which does not exist is passed in' do
       @worksheet.get_column_border_bottom(11).should be_nil
     end
   end

   describe '.get_column_border_diagonal' do
     it 'should return nil if no border is specified for that column in that direction' do
       @worksheet.get_column_border_diagonal(0).should be_nil
     end

     it 'should return type of border that this column has on diagonal' do
       @worksheet.change_column_border_diagonal(0,'thin')
       @worksheet.get_column_border_diagonal(0).should == 'thin'
     end

     it 'should cause error if a negative argument is passed in' do
       lambda {
         @worksheet.get_column_border_diagonal(-1)
       }.should raise_error
     end

     it 'should return nil if a column which does not exist is passed in' do
       @worksheet.get_column_border_diagonal(11).should be_nil
     end
   end

end
