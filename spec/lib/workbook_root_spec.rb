# frozen_string_literal: true

require 'spec_helper'

describe RubyXL::WorkbookRoot do
  def workbook_with_shared_drawing
    workbook = RubyXL::Workbook.new
    workbook.add_worksheet('Sheet2')

    workbook.worksheets.each { |sheet|
      image = RubyXL::BinaryImageFile.new(Pathname.new('/xl/media/image1.png'), 'PNG')
      drawing = RubyXL::DrawingFile.new(Pathname.new('/xl/drawings/drawing1.xml'), '<xdr:wsDr/>')
      drawing.generic_storage = [ image ]
      drawing.relationship_container = 
        RubyXL::OOXMLRelationshipsFile.new(:relationships => [
          RubyXL::Relationship.new(:id => 'rId1', :type => image.class::REL_TYPE, :target => '../media/image1.png')
        ])
      sheet.generic_storage = [ drawing ]
      sheet.relationship_container =
        RubyXL::OOXMLRelationshipsFile.new(:relationships => [
          RubyXL::Relationship.new(:id => 'rId1', :type => drawing.class::REL_TYPE, :target => '../drawings/drawing1.xml')
        ])
    }

    workbook
  end

  describe '#stream' do
    let(:binary_data) { workbook_with_shared_drawing.stream.read }

    it 'should keep the shared parts reachable from every sheet' do
      RubyXL::Parser.parse_buffer(binary_data).worksheets.each { |sheet| 
        drawing = sheet.generic_storage.first
        expect(drawing).to be_an_instance_of(RubyXL::DrawingFile)
        image = drawing.generic_storage.first
        expect(image).to be_an_instance_of(RubyXL::BinaryImageFile)

        expect(image.xlsx_path.to_s).to eq('/xl/media/image1.png')
        expect(image.data).to eq('PNG')
      }
    end
  end
end
