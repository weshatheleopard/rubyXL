module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_v-1.html
  class CellValue < OOXMLObject
    define_attribute(:_, :string, :accessor => :value)
    define_element_name 'v'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_c-2.html
  class Cell < OOXMLObject
    define_attribute(:r,   :ref)
    define_attribute(:s,   :integer)
    define_attribute(:t,   :string,  :default => 'n', :values => %w{ b n e s str inlineStr })
    define_attribute(:cm,  :integer)
    define_attribute(:vm,  :integer)
    define_attribute(:ph,  :bool)
#    define_child_node(RubyXL::CellFormula) # f 
    define_child_node(RubyXL::CellValue)
#    define_child_node(RubyXL::RichText)  # is
    define_element_name 'c'

    def row
      r && r.first_row
    end

    def row=(v)
      self.r = RubyXL::Reference.new(v, column || 0)
    end

    def column
      r && r.first_col
    end

    def column=(v)
      self.r = RubyXL::Reference.new(row || 0, v)
    end

    include LegacyCell
  end

#TODO#<row r="1" spans="1:1" x14ac:dyDescent="0.25">

  # http://www.schemacentral.com/sc/ooxml/e-ssml_row-1.html
  class Row < OOXMLObject
    define_attribute(:r,            :integer)
    define_attribute(:spans,        :string)
    define_attribute(:s,            :integer)
    define_attribute(:customFormat, :bool,    :default => false)
    define_attribute(:ht,           :float)
    define_attribute(:hidden,       :bool,    :default => false)
    define_attribute(:customHeight, :bool,    :default => false)
    define_attribute(:outlineLevel, :integer, :default => 0)
    define_attribute(:collapsed,    :bool,    :default => false)
    define_attribute(:thickTop,     :bool,    :default => false)
    define_attribute(:thickBot,     :bool,    :default => false)
    define_attribute(:ph,           :bool,    :default => false)
    define_child_node(RubyXL::Cell, :collection => true, :accessor => :cells)
    define_element_name 'row'

    def [](ind)
      cells[ind]
    end

  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sheetData-1.html
  class SheetData < OOXMLObject
    define_child_node(RubyXL::Row, :collection => true, :accessor => :rows)
    define_element_name 'sheetData'

    def [](ind)
      rows[ind]
    end

  end

end
