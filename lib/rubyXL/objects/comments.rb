require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/extensions'

module RubyXL

  # http://www.schemacentral.com/sc/ooxml/e-ssml_comment-1.html
  class Comment < OOXMLObject
    define_child_node(RubyXL::RichText, :node_name => 'text')
    define_attribute(:ref,      :ref, :required => true)
    define_attribute(:authorId, :int, :required => true)
    define_attribute(:guid,     :string)
    define_element_name 'comment'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_authors-1.html
  class CommentList < OOXMLContainerObject
    define_child_node(RubyXL::Comment)
    define_element_name 'commentList'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_authors-1.html
  class Authors < OOXMLContainerObject
    define_child_node(RubyXL::StringNode, :node_name => :author, :collection => :true)
    define_element_name 'authors'
  end

  # http://www.schemacentral.com/sc/ooxml/e-ssml_comments.html
  class Comments < OOXMLTopLevelObject
    define_child_node(RubyXL::Authors)
    define_child_node(RubyXL::CommentList)
    define_child_node(RubyXL::ExtensionStorageArea)
    define_element_name 'comments'
    set_namespaces('xmlns' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')

    def self.xlsx_path
      File.join('xl', 'comments1.xml')
    end

    def self.content_type
      'application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml'
    end

  end

end
