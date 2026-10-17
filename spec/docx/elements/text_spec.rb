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
  end
end
