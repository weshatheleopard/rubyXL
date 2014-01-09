module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_definedName-2.html
  class DefinedName < OOXMLObject
    define_attribute(:name,              :string, :required => true)
    define_attribute(:comment,           :string)
    define_attribute(:customMenu,        :string)
    define_attribute(:description,       :string)
    define_attribute(:help,              :string)
    define_attribute(:description,       :string)
    define_attribute(:localSheetId,      :string)

    define_attribute(:hidden,            :bool, :default => false)
    define_attribute(:function,          :bool, :default => false)
    define_attribute(:vbProcedure,       :bool, :default => false)
    define_attribute(:xlm,               :bool, :default => false)

    define_attribute(:functionGroupId,   :int)
    define_attribute(:shortcutKey,       :string)
    define_attribute(:publishToServer,   :bool, :default => false)
    define_attribute(:workbookParameter, :bool, :default => false)

    define_attribute(:_,                 :string, :accessor => :reference)
    define_element_name 'definedName'
  end

end
