# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Containers::Container do
  let(:container_class) do
    Class.new do
      include Docx::Elements::Containers::Container

      def initialize(node)
        @node = node
        @properties_tag = 'pPr'
      end
    end
  end

  let(:body) do
    w_body(<<~XML)
      <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>Hello </w:t></w:r><w:r><w:t>world</w:t></w:r></w:p>
      <w:p><w:r><w:t>other</w:t></w:r></w:p>
    XML
  end
  let(:node) { body.element_children.first }
  subject(:container) { container_class.new(node) }

  describe '#properties' do
    it 'returns the properties node named by @properties_tag' do
      expect(container.properties.name).to eq('pPr')
    end

    it 'returns nil when there is no properties node' do
      expect(container_class.new(body.element_children.last).properties).to be_nil
    end
  end

  describe '#blank!' do
    it 'empties every w:t while keeping the runs' do
      container.blank!

      expect(node.xpath('.//w:t').map(&:text)).to eq(['', ''])
    end
  end

  describe '#remove!' do
    it 'removes the node from its parent' do
      container.remove!

      expect(body.element_children.map(&:text)).to eq(['other'])
    end
  end
end
