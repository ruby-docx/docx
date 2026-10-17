# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Element do
  let(:element_class) do
    Class.new do
      include Docx::Elements::Element

      def self.tag
        'p'
      end

      def initialize(node)
        @node = node
      end
    end
  end

  let(:body) { w_body('<w:p><w:r><w:t>first</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>') }
  let(:first) { element_class.new(body.element_children[0]) }
  let(:second) { element_class.new(body.element_children[1]) }

  def texts
    body.element_children.map(&:text)
  end

  describe '.included' do
    it 'defines a TAG constant on the including class' do
      expect(element_class::TAG).to eq('')
    end
  end

  describe '#xpath / #at_xpath' do
    it 'delegates to the node' do
      expect(first.xpath('.//w:t').map(&:text)).to eq(['first'])
      expect(first.at_xpath('.//w:t').text).to eq('first')
    end
  end

  describe '#parent' do
    it 'returns the parent node' do
      expect(first.parent.name).to eq('body')
    end

    it 'returns nil when the parent is not of the given type' do
      expect(first.parent('w:tc')).to be_nil
    end
  end

  describe '#parent_paragraph' do
    it 'wraps the enclosing w:p in a Paragraph' do
      run = element_class.new(first.node.at_xpath('w:r'))

      paragraph = run.parent_paragraph

      expect(paragraph).to be_a(Docx::Elements::Containers::Paragraph)
      expect(paragraph.text).to eq('first')
    end
  end

  describe '#copy' do
    it 'returns a detached element of the same class with a duplicated node' do
      copy = first.copy

      expect(copy).to be_an_instance_of(element_class)
      expect(copy.node).not_to equal(first.node)
      expect(copy.node.text).to eq(first.node.text)
      expect(copy.node.parent).to be_nil
    end
  end

  describe '#insert_after' do
    it 'inserts the node as the next sibling and returns self' do
      copy = first.copy

      expect(copy.insert_after(first)).to equal(copy)
      expect(texts).to eq(%w[first first second])
    end
  end

  describe '#insert_before' do
    it 'inserts the node as the previous sibling' do
      second.copy.insert_before(first)

      expect(texts).to eq(%w[second first second])
    end
  end

  describe '#append_to' do
    it 'moves the node to be the last child of the given element' do
      run = element_class.new(second.node.at_xpath('w:r'))

      run.append_to(first)

      expect(texts).to eq(['firstsecond', ''])
    end
  end

  describe '.create_with' do
    it 'builds a detached element from the class tag in the same XML document' do
      created = element_class.create_with(first)

      expect(created.node.name).to eq('w:p')
      expect(created.node.document).to equal(body.document)
      expect(created.node.parent).to be_nil
    end
  end

  describe '.create_within' do
    it 'builds an element and appends it to the given element' do
      created = element_class.create_within(first)

      expect(first.node.element_children.last).to equal(created.node)
    end
  end

  describe '#html_tag' do
    it 'wraps content in the tag' do
      expect(first.html_tag(:p, content: 'hi')).to eq('<p>hi</p>')
    end

    it 'renders an empty tag without content' do
      expect(first.html_tag(:p)).to eq('<p></p>')
    end

    it 'renders styles as an inline style attribute' do
      html = first.html_tag(:span, content: 'x', styles: { 'font-size' => '12pt', 'color' => '#FF0000' })

      expect(html).to eq('<span style="font-size:12pt;color:#FF0000;">x</span>')
    end

    it 'renders attributes' do
      html = first.html_tag(:a, content: 'x', attributes: { href: 'http://example.com/', target: '_blank' })

      expect(html).to eq('<a href="http://example.com/" target="_blank">x</a>')
    end

    it 'ignores empty styles and attributes' do
      expect(first.html_tag(:p, content: 'x', styles: {}, attributes: {})).to eq('<p>x</p>')
    end
  end
end
