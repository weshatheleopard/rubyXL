require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/simple_types'
require 'rubyXL/objects/extensions'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_dateGroupItem-1.html
  class DateGroupItem < OOXMLObject
    define_attribute(:year,   :int, :required => true)
    define_attribute(:month,  :int)
    define_attribute(:day,    :int)
    define_attribute(:hour,   :int)
    define_attribute(:minute, :int)
    define_attribute(:second, :int)
    define_attribute(:dateTimeGrouping, :string, :values => RubyXL::ST_DateTimeGrouping)
    define_element_name 'dateGroupItem'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_filters-1.html
  class FilterContainer < OOXMLObject
    define_attribute(:blank,        :bool,  :default  => false)
    define_attribute(:calendarType, :string, :default => 'none', :values => RubyXL::ST_CalendarType)
    define_child_node(RubyXL::StringValue,    :node_name => :filter, :collection => true, :accessor => :filters)
    define_child_node(RubyXL::DateGroupItem, :collection => true, :accessor => :date_group_items)
    define_element_name 'filters'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_top10-1.html
  class Top10 < OOXMLObject
    define_attribute(:top,       :bool,  :default  => true)	
    define_attribute(:percent,   :bool,  :default  => false)
    define_attribute(:val,       :float, :required => true)
    define_attribute(:filterVal, :float)
    define_element_name 'top10'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_customFilter-1.html
  class CustomFilter < OOXMLObject
    define_attribute(:operator, :string, :default => 'equal', :values => RubyXL::ST_FilterOperator)
    define_attribute(:val, :string)
    define_element_name 'customFilter'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_customFilters-1.html
  class CustomFilterContainer < OOXMLObject
    define_attribute(:and, :bool,  :default  => false)
    define_child_node(RubyXL::CustomFilter, :collection => true, :accessor => :custom_filters)
    define_element_name 'customFilters'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_dynamicFilter-1.html
  class DynamicFilter < OOXMLObject
    define_attribute(:type,   :string, :required => true, :values => RubyXL::ST_DynamicFilterType)
    define_attribute(:val,    :float)
    define_attribute(:maxVal, :float)
    define_element_name 'dynamicFilter'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_colorFilter-1.html
  class ColorFilter < OOXMLObject
    define_attribute(:dxfId,     :string)
    define_attribute(:cellColor, :bool)
    define_element_name 'colorFilter'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_iconFilter-1.html
  class IconFilter < OOXMLObject
    define_attribute(:iconSet, :string, :values => RubyXL::ST_IconSetType)
    define_attribute(:iconId,  :int)
    define_element_name 'iconFilter'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_filterColumn-1.html
  class AutoFilterColumn < OOXMLObject
    define_attribute(:colId,        :int,  :required => true)
    define_attribute(:hiddenButton, :bool, :default  => false)
    define_attribute(:showButton,   :bool, :default  => true)	
    define_child_node(RubyXL::FilterContainer)
    define_child_node(RubyXL::Top10)
    define_child_node(RubyXL::CustomFilterContainer, :accessor => :custom_filter_container)
    define_child_node(RubyXL::DynamicFilter)
    define_child_node(RubyXL::ColorFilter)
    define_child_node(RubyXL::IconFilter)
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'filterColumn'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sortCondition-1.html
  class SortCondition < OOXMLObject
    define_attribute(:descending, :bool,   :default  => false)
    define_attribute(:sortBy,     :string, :default => 'value', :values => RubyXL::ST_SortBy)
    define_attribute(:ref,        :ref,    :required => true)
    define_attribute(:customList, :string)
    define_attribute(:dxfId,      :int)
    define_attribute(:iconSet,    :string, :required => true, :default => '3Arrows',
                       :values => RubyXL::ST_IconSetType)
    define_attribute(:iconId,     :int)
    define_element_name 'sortCondition'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_sortState-2.html
  class SortState < OOXMLObject
    define_attribute(:columnSort,    :bool,   :default  => false)
    define_attribute(:caseSensitive, :bool,   :default  => false)
    define_attribute(:sortMethod,    :string, :default => 'none', :values => RubyXL::ST_SortMethod)
    define_attribute(:ref,           :ref,    :required => true)
    define_child_node(RubyXL::SortCondition,  :colection => true)
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'sortState'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_autoFilter-2.html
  class AutoFilter < OOXMLObject
    define_attribute(:ref, :ref)
    define_child_node(RubyXL::AutoFilterColumn)
    define_child_node(RubyXL::SortState)
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'autoFilter'
  end

end