require 'rubyXL/objects/root'
require 'rubyXL/parser'

module RubyXL

  def self.from_root(path)
    return path unless path.absolute?
    path.relative_path_from(OOXMLTopLevelObject::ROOT)
  end

end
