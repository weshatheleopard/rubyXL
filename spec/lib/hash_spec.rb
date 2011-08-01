require 'rubygems'
require 'rubyXL'

describe RubyXL::Hash do
  before do
    @xml = '<root xmlns:foo="bar"><bar hello="world"/></root>'
    @hash = {
              :root => {
                         :bar => { :attributes => { :hello => 'world' } }
                       }
            }
  end

  describe '.from_xml' do
    it 'should create a hash which correctly corresponds to XML' do
      nokogiri = RubyXL::Hash.from_xml(@xml)
      nokogiri.should == @hash
    end
  end

  describe '.xml_node_to_hash' do
    it 'should create a hash which correctly corresponds to a Nokogiri root node' do
      nokogiri = Nokogiri::XML::Document.parse(@xml)
      my_hash = RubyXL::Hash.xml_node_to_hash(nokogiri.root)
      my_hash.should == @hash[:root]
    end
  end
end
