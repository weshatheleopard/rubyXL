module RubyXL
  # TODO: http://www.schemacentral.com/sc/ooxml/e-ssml_definedName-2.html

  class DefinedName < OOXMLObject
    define_attribute(:name,              :name,              :string, :required)
    define_attribute(:comment,           :comment,           :string)
    define_attribute(:custom_menu,       :customMenu,        :string)
    define_attribute(:description,       :description,       :string)
    define_attribute(:help,              :help,              :string)
    define_attribute(:description,       :description,       :string)
    define_attribute(:local_sheet_id,    :localSheetId,      :string)

    define_attribute(:hidden,            :hidden,            :bool,   false, false)
    define_attribute(:function,          :function,          :bool,   false, false)
    define_attribute(:vb_procedure,      :vbProcedure,       :bool,   false, false)
    define_attribute(:xlm,               :xlm,               :bool,   false, false)

    define_attribute(:function_group_id, :functionGroupId,   :int)
    define_attribute(:shortcut_key,      :shortcutKey,       :string)
    define_attribute(:publish_to_server, :publishToServer,   :bool,   false, false)
    define_attribute(:workbookParameter, :workbookParameter, :bool,   false, false)

    define_attribute(:reference,         :_,                 :string)
    define_element_name 'definedName'
  end

end
