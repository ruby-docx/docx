# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Containers::TableColumn do
  let(:cell_nodes) do
    w_node(<<~XML).xpath('w:tc')
      <w:tr>
        <w:tc><w:p><w:r><w:t>aphids</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>puceron</w:t></w:r></w:p></w:tc>
      </w:tr>
    XML
  end
  subject(:column) { described_class.new(cell_nodes, font_size: 11) }

  it 'is backed by the w:gridCol tag' do
    expect(described_class.tag).to eq('w:gridCol')
  end

  describe '#cells' do
    it 'wraps the given cell nodes as TableCells' do
      expect(column.cells).to all(be_an_instance_of(Docx::Elements::Containers::TableCell))
      expect(column.cells.map(&:text)).to eq(%w[aphids puceron])
    end

    it 'passes the document properties on to the cells' do
      expect(column.cells.first.paragraphs.first.font_size).to eq(11)
    end
  end
end
