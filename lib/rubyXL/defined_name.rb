module RubyXL

  class DefinedName
    attr_accessor :name, :reference

    def initialize(args = {})
    end

    def self.parse(xml)
      defined_name = self.new

#<definedNames>
#<definedName name="CoolName1">'Defined Names Test'!$A$1</definedName>
#<definedName name="NotSoCoolName2">'Defined Names Test'!$A$2</definedName>
#</definedNames>
puts xml.inspect

      defined_name.name = xml.attributes['name'].value
      defined_name.reference = xml.text
      defined_name
    end 

    def write_xml(xml)
      xml << xml.create_element('definedName', { :name => name }, reference)
    end

  end


end
