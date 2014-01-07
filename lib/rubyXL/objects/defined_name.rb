module RubyXL
  # TODO: http://www.schemacentral.com/sc/ooxml/e-ssml_definedName-2.html

  class DefinedName < OOXMLObject
    define_attribute(:name,              :name,              :string)
    define_attribute(:comment,           :comment,           :string, true)
    define_attribute(:custom_menu,       :customMenu,        :string, true)
    define_attribute(:description,       :description,       :string, true)
    define_attribute(:help,              :help,              :string, true)
    define_attribute(:description,       :description,       :string, true)
    define_attribute(:local_sheet_id,    :localSheetId,      :string, true)

    define_attribute(:hidden,            :hidden,            :bool,   true, false)
    define_attribute(:function,          :function,          :bool,   true, false)
    define_attribute(:vb_procedure,      :vbProcedure,       :bool,   true, false)
    define_attribute(:xlm,               :xlm,               :bool,   true, false)

    define_attribute(:function_group_id, :functionGroupId,   :int,    true)
    define_attribute(:shortcut_key,      :shortcutKey,       :string, true)
    define_attribute(:publish_to_server, :publishToServer,   :bool,   true, false)
    define_attribute(:workbookParameter, :workbookParameter, :bool,   true, false)

    define_attribute(:reference,         :_,                 :string)
    define_element_name 'definedName'
  end

end
