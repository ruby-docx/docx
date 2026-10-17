# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Containers::TableCell do
  let(:properties) { { font_size: 11, hyperlinks: { 'rId1' => 'http://google.com/' } } }
  let(:node) do
    w_node(<<~XML)
      <w:tc>
        <w:tcPr/>
        <w:p><w:r><w:t>Directive </w:t></w:r></w:p>
        <w:p><w:hyperlink r:id="rId1"><w:r><w:t>link</w:t></w:r></w:hyperlink></w:p>
      </w:tc>
    XML
  end
  subject(:cell) { described_class.new(node, properties) }

  it 'is backed by the w:tc tag' do
    expect(described_class.tag).to eq('tc')
  end

  describe '#paragraphs' do
    it 'wraps each w:p as a Paragraph' do
      expect(cell.paragraphs).to all(be_an_instance_of(Docx::Elements::Containers::Paragraph))
      expect(cell.paragraphs.size).to eq(2)
    end

    # Regression test for #160: rendering html for a cell that contains a
    # hyperlink must not raise (the cell's paragraphs need document_properties).
    it 'passes the document properties on so hyperlinks render as html' do
      expect(cell.paragraphs.map(&:to_html).join).to include('<a href="http://google.com/" target="_blank">link</a>')
    end
  end

  describe '#text' do
    it 'joins the text of its paragraphs' do
      expect(cell.text).to eq('Directive link')
      expect(cell.to_s).to eq('Directive link')
    end
  end

  describe '#each_paragraph' do
    it 'yields each paragraph' do
      expect { |block| cell.each_paragraph(&block) }.to yield_control.twice
    end
  end
end
