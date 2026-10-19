# frozen_string_literal: true

require 'spec_helper'

describe RubyXL::WorkbookRoot do
  def relationships(obj, target)
    RubyXL::OOXMLRelationshipsFile.new(:relationships => [
      RubyXL::Relationship.new(:id => 'rId1', :type => obj.class::REL_TYPE, :target => target)
    ])
  end

  # The parser creates one object per relationship, so a part reached from several places (here one drawing
  # with its image, used by both sheets) is held in memory several times under the same path.
  def workbook_with_shared_drawing
    workbook = RubyXL::Workbook.new
    workbook.add_worksheet('Sheet2')

    workbook.worksheets.each { |sheet|
      image = RubyXL::BinaryImageFile.new(Pathname.new('/xl/media/image1.png'), 'PNG')
      drawing = RubyXL::DrawingFile.new(Pathname.new('/xl/drawings/drawing1.xml'), '<xdr:wsDr/>')
      drawing.generic_storage = [ image ]
      drawing.relationship_container = relationships(image, '../media/image1.png')
      sheet.generic_storage = [ drawing ]
      sheet.relationship_container = relationships(drawing, '../drawings/drawing1.xml')
    }

    workbook
  end

  describe '#stream' do
    let(:bytes) { workbook_with_shared_drawing.stream.string }

    it 'should store a part reached through several relationships only once' do
      # The central directory lists each name once however often it was written,
      # while a sequential read stops at the orphaned duplicate.
      central = Zip::File.open_buffer(bytes).map(&:name)
      sequential = []
      Zip::InputStream.open(StringIO.new(bytes)) { |zip| while (entry = zip.get_next_entry) do sequential << entry.name end }

      expect(central).to include('xl/drawings/drawing1.xml', 'xl/drawings/_rels/drawing1.xml.rels')
      expect(central).to include('xl/media/image1.png')
      expect(sequential).to eq(central)
    end

    it 'should keep the shared parts reachable from every sheet' do
      images = RubyXL::Parser.parse_buffer(bytes).worksheets.map { |sheet| sheet.generic_storage.first.generic_storage.first }

      expect(images.map { |image| image.xlsx_path.to_s }).to eq([ '/xl/media/image1.png' ] * 2)
      expect(images.map(&:data)).to eq([ 'PNG' ] * 2)
    end
  end
end
