# frozen_string_literal: true

require 'spec_helper'
require 'byebug'

describe RubyXL::DocumentPropertiesFile do

  it 'adds default namespaces if not present in original document' do
    original_xml = <<~EoT
    <?xml version="1.0" encoding="utf-8"?>
    <Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
      <Application>DevExpress Office File API/24.1.3.0</Application>
      <AppVersion>24.1</AppVersion>
    </Properties>
    EoT

    zip_entry = instance_double(Zip::Entry)
    allow(zip_entry).to receive(:get_input_stream).and_yield(original_xml)

    zip_file = instance_double(Zip::File, find_entry: zip_entry)

    properties = RubyXL::DocumentPropertiesFile.parse_file(zip_file, Pathname.new("/docProps/app.xml"))

    properties.root = RubyXL::WorkbookRoot.new
    properties.root.workbook = RubyXL::Workbook.new

    changed_xml = properties.write_xml

    node = Nokogiri::XML(changed_xml)
    expect(node.namespaces).to eql({
      "xmlns" => "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties",
      "xmlns:vt" => "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
    })
    expect(node.errors).to be_empty
  end

end
