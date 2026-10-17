# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Containers::StylesConfiguration do
  let(:raw_styles) do
    Nokogiri::XML(<<~XML)
      <w:styles xmlns:w="#{XmlHelpers::NAMESPACES['w']}">
        <w:docDefaults/>
        <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
        <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/></w:style>
      </w:styles>
    XML
  end
  subject(:configuration) { described_class.new(raw_styles) }

  describe '#styles' do
    it 'wraps each w:style with a style id, skipping other children' do
      expect(configuration.styles).to all(be_a(Docx::Elements::Style))
      expect(configuration.styles.map(&:id)).to eq(%w[Normal Heading1])
    end
  end

  describe '#size' do
    it 'counts the styles' do
      expect(configuration.size).to eq(2)
    end
  end

  describe '#style_of' do
    it 'finds a style by id or by name' do
      expect(configuration.style_of('Heading1').name).to eq('Heading 1')
      expect(configuration.style_of('Heading 1').id).to eq('Heading1')
    end

    it 'raises StyleNotFound for an unknown style' do
      expect { configuration.style_of('Red') }
        .to raise_error(Docx::Errors::StyleNotFound, "Style name or id 'Red' not found")
    end
  end

  describe '#add_style' do
    it 'adds a style named after its id by default' do
      red = configuration.add_style('Red')

      expect(configuration.size).to eq(3)
      expect(red).to be_a(Docx::Elements::Style)
      expect(red.id).to eq('Red')
      expect(red.name).to eq('Red')
      expect(configuration.style_of('Red').id).to eq('Red')
    end

    it 'applies the given attributes' do
      red = configuration.add_style('Red', name: 'Red text', font_color: 'FF0000', font_size: 20)

      expect(red.name).to eq('Red text')
      expect(red.font_color).to eq('FF0000')
      expect(red.font_size).to eq(20)
    end

    it 'validates attributes set afterwards' do
      red = configuration.add_style('Red')

      expect { red.font_color = '#FFFFFF' }.to raise_error(Docx::Errors::StyleInvalidPropertyValue)
      expect { red.font_color = 'blue' }.to raise_error(Docx::Errors::StyleInvalidPropertyValue)
      expect { red.font_color = 'FF0000' }.not_to raise_error
    end
  end

  describe '#remove_style' do
    it 'removes the style node' do
      configuration.remove_style('Heading1')

      expect(configuration.size).to eq(1)
      expect { configuration.style_of('Heading1') }.to raise_error(Docx::Errors::StyleNotFound)
    end
  end

  describe 'editing an existing style' do
    it 'writes changes through to the underlying node' do
      normal = configuration.style_of('Normal')
      expect(normal.font_color).to be_nil

      normal.font_color = '000000'

      expect(normal.font_color).to eq('000000')
      expect(normal.node.at_xpath('w:rPr/w:color/@w:val').value).to eq('000000')
    end
  end

  describe '#serialize' do
    it 'serializes the underlying styles document, including changes' do
      configuration.add_style('Red')

      expect(configuration.serialize).to include('w:styleId="Red"')
    end
  end
end
