# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Containers::TableRow do
  let(:node) do
    w_node(<<~XML)
      <w:tr>
        <w:tc><w:p><w:r><w:t>Second table</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Second tableau</w:t></w:r></w:p></w:tc>
      </w:tr>
    XML
  end
  subject(:row) { described_class.new(node, font_size: 11) }

  it 'is backed by the w:tr tag' do
    expect(described_class.tag).to eq('tr')
  end

  describe '#cells' do
    it 'wraps each w:tc as a TableCell' do
      expect(row.cells).to all(be_an_instance_of(Docx::Elements::Containers::TableCell))
      expect(row.cells.map(&:text)).to eq(['Second table', 'Second tableau'])
    end

    it 'passes the document properties on to the cells' do
      expect(row.cells.first.paragraphs.first.font_size).to eq(11)
    end
  end
end
