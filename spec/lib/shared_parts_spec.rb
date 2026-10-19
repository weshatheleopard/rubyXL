# frozen_string_literal: true

require 'spec_helper'
require 'zlib'

describe RubyXL::WorkbookRoot do
  REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

  def relationships(type, target)
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' \
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' \
      "<Relationship Id=\"rId1\" Type=\"#{REL_NS}/#{type}\" Target=\"#{target}\"/></Relationships>"
  end

  def png_chunk(type, data)
    [data.bytesize].pack('N') + type + data + [Zlib.crc32(type + data)].pack('N')
  end

  # Two worksheets, each with a drawing; both drawings embed xl/media/image1.png.
  def workbook_with_shared_image
    workbook = RubyXL::Workbook.new
    workbook.add_worksheet('Sheet2')
    workbook.worksheets.each { |sheet| sheet.add_cell(0, 0, sheet.sheet_name) }

    entries = {}
    Zip::File.open_buffer(workbook.stream.string) { |zip| zip.each { |e| entries[e.name] = e.get_input_stream.read } }

    drawing = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' \
      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" ' \
      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"#{REL_NS}\">" \
      '<xdr:oneCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' \
      '<xdr:ext cx="9525" cy="9525"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="2" name="Logo"/><xdr:cNvPicPr/></xdr:nvPicPr>' \
      '<xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' \
      '<xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor></xdr:wsDr>'

    [1, 2].each { |n|
      sheet = entries["xl/worksheets/sheet#{n}.xml"]
      sheet = sheet.sub('<worksheet ', "<worksheet xmlns:r=\"#{REL_NS}\" ") unless sheet.include?('xmlns:r=')
      entries["xl/worksheets/sheet#{n}.xml"] = sheet.sub('</worksheet>', '<drawing r:id="rId1"/></worksheet>')
      entries["xl/worksheets/_rels/sheet#{n}.xml.rels"] = relationships('drawing', "../drawings/drawing#{n}.xml")
      entries["xl/drawings/drawing#{n}.xml"] = drawing
      entries["xl/drawings/_rels/drawing#{n}.xml.rels"] = relationships('image', '../media/image1.png')
    }

    entries['xl/media/image1.png'] = "\x89PNG\r\n\x1a\n".b +
      png_chunk('IHDR', [1, 1, 8, 2, 0, 0, 0].pack('NNCCCCC')) +
      png_chunk('IDAT', Zlib::Deflate.deflate("\x00\xff\x00\x00".b)) +
      png_chunk('IEND', '')

    entries['[Content_Types].xml'] = entries['[Content_Types].xml'].sub('</Types>',
      '<Default Extension="png" ContentType="image/png"/>' +
      [1, 2].map { |n| "<Override PartName=\"/xl/drawings/drawing#{n}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>" }.join +
      '</Types>')

    Zip::OutputStream.write_buffer(suppress_extra_fields: true) { |zos| entries.each { |name, data| zos.put_next_entry(name); zos.write(data) } }.string
  end

  # Walks the local file headers, which is where a duplicate lives: the central
  # directory lists each name once. Sizes are in the headers (no data descriptors).
  def stored_entry_names(bytes)
    names = []
    offset = 0
    while bytes.byteslice(offset, 4) == "PK\x03\x04".b
      compressed_size, = bytes.byteslice(offset + 18, 4).unpack('V')
      name_length, extra_length = bytes.byteslice(offset + 26, 4).unpack('vv')
      names << bytes.byteslice(offset + 30, name_length)
      offset += 30 + name_length + extra_length + compressed_size
    end
    names
  end

  describe '#stream' do
    let(:source) { workbook_with_shared_image }

    it 'should store a part shared by several relationships only once' do
      expect(stored_entry_names(source).count('xl/media/image1.png')).to eq(1)

      bytes = RubyXL::Parser.parse_buffer(source).stream.string
      names = stored_entry_names(bytes)
      central_directory_size = nil
      Zip::File.open_buffer(bytes) { |zip| central_directory_size = zip.size }

      expect(names.count('xl/media/image1.png')).to eq(1)
      expect(names).to eq(names.uniq)
      expect(names.size).to eq(central_directory_size)
    end

    it 'should keep the shared part reachable from every drawing' do
      reparsed = RubyXL::Parser.parse_buffer(RubyXL::Parser.parse_buffer(source).stream.string)

      images = reparsed.worksheets.map { |sheet|
        drawing = sheet.relationship_container.related_files.values.grep(RubyXL::DrawingFile).first
        drawing.relationship_container.related_files.values.first
      }

      expect(images.map { |image| image.xlsx_path.to_s }).to eq(['/xl/media/image1.png', '/xl/media/image1.png'])
      expect(images.map(&:data).uniq.size).to eq(1)
    end
  end
end
