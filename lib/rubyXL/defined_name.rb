module RubyXL

  class DefinedName
    attr_accessor :name, :reference

    def initialize(args = {})
    end

    def self.parse(xml)
      defined_name = self.new
      defined_name.name = xml.attributes['name'].value
      defined_name.reference = xml.text
      defined_name
    end 

    def write_xml(xml)
      xml.create_element('definedName', { :name => name }, reference)
    end

  end

end
