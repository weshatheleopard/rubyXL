module RubyXL

  # http://www.datypic.com/sc/ooxml/e-ssml_sheetName-1.html
  class SheetName < OOXMLObject
    define_attribute(:val, :string)
    define_element_name 'sheetName'
  end

  # http://www.datypic.com/sc/ooxml/e-ssml_sheetNames-1.html
  class SheetNames < OOXMLContainerObject
    define_child_node(RubyXL::SheetName, :collection => true)
    define_element_name 'sheetNames'
  end

  # http://www.datypic.com/sc/ooxml/e-ssml_definedName-1.html
  class DefinedNameExt < OOXMLObject
    define_attribute(:name,     :string, :required => true)
    define_attribute(:refersTo, :string)
    define_attribute(:sheetId,  :uint)
    define_element_name 'definedName'
  end

  # http://www.datypic.com/sc/ooxml/e-ssml_definedNames-1.html
  class DefinedNamesExt < OOXMLContainerObject
    define_child_node(RubyXL::DefinedNameExt, :collection => true)
    define_element_name 'definedNames'
  end

  # http://www.datypic.com/sc/ooxml/e-ssml_cell-1.html
  class CellExt < OOXMLObject
    define_child_node(RubyXL::StringNode, :node_name => :v)
    define_attribute(:r,  :sqref)
    define_attribute(:t,  :string, :default => 'n')
    define_attribute(:vm, :uint,   :default => 0)
    define_element_name 'cell'
  end

  # http://www.datypic.com/sc/ooxml/e-ssml_row-2.html
  class RowExt < OOXMLObject
    define_child_node(RubyXL::CellExt, :collection => true)
    define_attribute(:r, :uint, :required => true)
    define_element_name 'row'
  end

  # http://www.datypic.com/sc/ooxml/e-ssml_sheetData-3.html
  class SheetDataExt < OOXMLObject
    define_child_node(RubyXL::RowExt, :collection => true)
    define_attribute(:sheetId,      :uint, :required => true)
    define_attribute(:refreshError, :bool)
    define_element_name 'sheetData'
  end

  # http://www.datypic.com/sc/ooxml/e-ssml_sheetDataSet-1.html
  class SheetDataSet < OOXMLContainerObject
    define_child_node(RubyXL::SheetDataExt, :collection => true)
    define_element_name 'sheetDataSet'
  end

  # http://www.datypic.com/sc/ooxml/e-ssml_externalBook-1.html
  class ExternalBook < OOXMLObject
    define_child_node(RubyXL::SheetNames)
    define_child_node(RubyXL::DefinedNamesExt)
    define_child_node(RubyXL::SheetDataSet)
    define_attribute(:'r:id', :string, :required => true)
    define_element_name 'externalBook'
  end

  class ExternalLinksFile < OOXMLTopLevelObject
    CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml'
    REL_TYPE     = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink'

    include RubyXL::RelationshipSupport
    define_child_node(RubyXL::ExternalBook)

    define_element_name 'externalLink'
    set_namespaces('http://schemas.openxmlformats.org/spreadsheetml/2006/main' => nil,
                   'http://schemas.openxmlformats.org/markup-compatibility/2006' => 'mc',
                   'http://schemas.openxmlformats.org/officeDocument/2006/relationships' => 'r')

    def xlsx_path
      ROOT.join('xl', 'externalLinks', "externalLink#{file_index}.xml")
    end
  end

end
