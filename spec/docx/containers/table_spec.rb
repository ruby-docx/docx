# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Containers::Table do
  let(:document) { instance_double(Docx::Document) }
  let(:properties) { { font_size: 11 } }
  let(:node) do
    w_node(<<~XML)
      <w:tbl>
        <w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>
        <w:tr>
          <w:tc><w:p><w:r><w:t>ENGLISH</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>FRANÇAIS</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>aphids</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>puceron</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t>bee</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t>abeille</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    XML
  end
  subject(:table) { described_class.new(node, properties, document) }

  it 'is backed by the w:tbl tag' do
    expect(described_class.tag).to eq('tbl')
  end

  describe '#rows' do
    it 'wraps each w:tr as a TableRow' do
      expect(table.rows).to all(be_an_instance_of(Docx::Elements::Containers::TableRow))
      expect(table.rows.map { |row| row.cells.map(&:text) }).to eq(
        [%w[ENGLISH FRANÇAIS], %w[aphids puceron], %w[bee abeille]]
      )
    end

    it 'passes the document properties and document on to the cells' do
      paragraph = table.rows[0].cells[0].paragraphs.first

      expect(paragraph.font_size).to eq(11)
      expect(paragraph.instance_variable_get(:@document)).to equal(document)
    end
  end

  describe '#row_count' do
    it 'counts the w:tr' do
      expect(table.row_count).to eq(3)
    end
  end

  describe '#columns' do
    it 'builds one TableColumn per grid column' do
      expect(table.columns).to all(be_an_instance_of(Docx::Elements::Containers::TableColumn))
      expect(table.columns.map { |column| column.cells.map(&:text) }).to eq(
        [%w[ENGLISH aphids bee], %w[FRANÇAIS puceron abeille]]
      )
    end
  end

  describe '#column_count' do
    it 'counts the w:gridCol' do
      expect(table.column_count).to eq(2)
    end
  end

  describe '#each_rows' do
    it 'yields each row' do
      expect { |block| table.each_rows(&block) }.to yield_control.exactly(3).times
    end
  end
end
