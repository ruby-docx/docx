# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Text do
  let(:node) { w_node('<w:t>hello</w:t>') }
  subject(:text) { described_class.new(node) }

  it 'is backed by the w:t tag' do
    expect(described_class.tag).to eq('t')
  end

  describe '#content' do
    it 'returns the node content' do
      expect(text.content).to eq('hello')
    end
  end

  describe '#content=' do
    it 'replaces the node content' do
      text.content = 'world'

      expect(node.content).to eq('world')
    end

    # Regression tests for #187: without xml:space="preserve", Word drops
    # leading and trailing spaces of a w:t.
    it 'marks the node xml:space="preserve" when the content has a trailing space' do
      text.content = 'Jane '

      expect(node.to_xml).to eq('<w:t xml:space="preserve">Jane </w:t>')
    end

    it 'marks the node xml:space="preserve" when the content has a leading space' do
      text.content = ' end'

      expect(node.to_xml).to eq('<w:t xml:space="preserve"> end</w:t>')
    end

    it 'does not add xml:space when the content has no surrounding whitespace' do
      text.content = 'Jane'

      expect(node.to_xml).to eq('<w:t>Jane</w:t>')
    end

    it 'keeps an existing xml:space="preserve"' do
      node = w_node('<w:t xml:space="preserve"> a </w:t>')

      described_class.new(node).content = 'b'

      expect(node.to_xml).to eq('<w:t xml:space="preserve">b</w:t>')
    end
  end
end
