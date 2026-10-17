# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Containers::Paragraph do
  def paragraph(xml, document_properties = {}, document = nil)
    described_class.new(w_node(xml), document_properties, document)
  end

  it 'is backed by the w:p tag' do
    expect(described_class.tag).to eq('p')
  end

  describe '#text' do
    it 'joins the text of its runs and hyperlinks' do
      p = paragraph(<<~XML)
        <w:p>
          <w:r><w:t>This sentence has a </w:t></w:r>
          <w:hyperlink r:id="rId1"><w:r><w:t>hyperlink</w:t></w:r></w:hyperlink>
          <w:r><w:t>.</w:t></w:r>
        </w:p>
      XML

      expect(p.text).to eq('This sentence has a hyperlink.')
      expect(p.to_s).to eq(p.text)
    end

    it 'is empty for a paragraph without runs' do
      expect(paragraph('<w:p/>').text).to eq('')
    end
  end

  describe '#text=' do
    it 'replaces the text of a single run, keeping its formatting' do
      p = paragraph('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>test text</w:t></w:r></w:p>')

      p.text = 'the real test'

      expect(p.text).to eq('the real test')
      expect(p.text_runs.first).to be_bolded
    end

    it 'adds a run when the paragraph has none' do
      p = paragraph('<w:p><w:pPr/></w:p>')

      p.text = 'new'

      expect(p.text).to eq('new')
    end

    it 'collapses multiple runs into one' do
      p = paragraph('<w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>')

      p.text = 'new'

      expect(p.text_runs.map(&:text)).to eq(['new'])
    end
  end

  describe '#text_runs' do
    it 'wraps w:r and w:hyperlink children as TextRuns' do
      p = paragraph('<w:p><w:pPr/><w:r><w:t>a</w:t></w:r><w:hyperlink r:id="rId1"><w:r><w:t>b</w:t></w:r></w:hyperlink></w:p>')

      expect(p.text_runs).to all(be_an_instance_of(Docx::Elements::Containers::TextRun))
      expect(p.text_runs.map(&:text)).to eq(%w[a b])
    end

    it 'passes the document properties on to its runs' do
      p = paragraph('<w:p><w:hyperlink r:id="rId1"><w:r><w:t>b</w:t></w:r></w:hyperlink></w:p>',
                    hyperlinks: { 'rId1' => 'http://example.com/' })

      expect(p.text_runs.first.href).to eq('http://example.com/')
    end
  end

  describe '#each_text_run' do
    it 'yields each text run' do
      p = paragraph('<w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>b</w:t></w:r></w:p>')

      expect { |block| p.each_text_run(&block) }.to yield_successive_args(
        an_object_having_attributes(text: 'a'), an_object_having_attributes(text: 'b')
      )
    end
  end

  # Regression tests for #147: Word often splits a placeholder across runs.
  describe '#substitute' do
    let(:split) do
      paragraph(<<~XML)
        <w:p>
          <w:r><w:t xml:space="preserve">Hello </w:t></w:r>
          <w:r><w:rPr><w:b/></w:rPr><w:t>{{fi</w:t></w:r>
          <w:r><w:t>rst_na</w:t></w:r>
          <w:r><w:t>me}}</w:t></w:r>
          <w:r><w:t>!</w:t></w:r>
        </w:p>
      XML
    end

    it 'replaces a placeholder split across runs, preserving surrounding text' do
      split.substitute('{{first_name}}', 'World')

      expect(split.text).to eq('Hello World!')
    end

    it 'collapses the match into the first run it touches, keeping that run formatting' do
      split.substitute('{{first_name}}', 'World')

      expect(split.text_runs.map(&:text)).to eq(['Hello ', 'World', '', '', '!'])
      expect(split.text_runs[1]).to be_bolded
    end

    it 'supports regex patterns and capture groups' do
      split.substitute(/\{\{(\w+)\}\}/, 'name=\1')

      expect(split.text).to eq('Hello name=first_name!')
    end

    it 'replaces every occurrence' do
      p = paragraph('<w:p><w:r><w:t>{{a}} and {{</w:t></w:r><w:r><w:t>a}}</w:t></w:r></w:p>')

      p.substitute('{{a}}', 'x')

      expect(p.text).to eq('x and x')
    end

    it 'does not re-match inside the inserted replacement' do
      p = paragraph('<w:p><w:r><w:t>{{a}}</w:t></w:r></w:p>')

      p.substitute('{{a}}', '{{a}}{{a}}')

      expect(p.text).to eq('{{a}}{{a}}')
    end

    it 'leaves the paragraph untouched when nothing matches' do
      expect(split.substitute('{{missing}}', 'x').text).to eq('Hello {{first_name}}!')
    end

    it 'returns self' do
      expect(split.substitute('{{first_name}}', 'World')).to equal(split)
    end
  end

  describe 'alignment' do
    it 'is centered with w:jc="center"' do
      p = paragraph('<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>')

      expect(p).to be_aligned_center
      expect(p).not_to be_aligned_left
      expect(p).not_to be_aligned_right
    end

    it 'is right aligned with w:jc="right"' do
      p = paragraph('<w:p><w:pPr><w:jc w:val="right"/></w:pPr></w:p>')

      expect(p).to be_aligned_right
      expect(p).not_to be_aligned_left
      expect(p).not_to be_aligned_center
    end

    it 'is left aligned with w:jc="left" or without w:jc' do
      expect(paragraph('<w:p><w:pPr><w:jc w:val="left"/></w:pPr></w:p>')).to be_aligned_left
      expect(paragraph('<w:p/>')).to be_aligned_left
      expect(paragraph('<w:p/>')).not_to be_aligned_center
    end
  end

  # ECMA-376 17.3.2.38: w:sz is in half-points, so 14pt is stored as 28.
  describe '#font_size' do
    it 'reads the paragraph size in points' do
      expect(paragraph('<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr></w:p>', font_size: 11).font_size).to eq(14)
    end

    it 'falls back to the document font size' do
      expect(paragraph('<w:p/>', font_size: 11).font_size).to eq(11)
    end
  end

  describe '#font_color' do
    it 'reads the color of its runs' do
      expect(paragraph('<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>x</w:t></w:r></w:p>').font_color).to eq('FF0000')
    end

    it 'is nil without a color' do
      expect(paragraph('<w:p><w:r><w:t>x</w:t></w:r></w:p>').font_color).to be_nil
    end

    # Regression test for #127
    it 'is nil when the color tag has no value' do
      expect(paragraph('<w:p><w:r><w:rPr><w:color/></w:rPr><w:t>x</w:t></w:r></w:p>').font_color).to be_nil
    end
  end

  describe '#to_html' do
    it 'wraps its runs in a p tag with the font size' do
      p = paragraph('<w:p><w:r><w:t>Normal</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r></w:p>', font_size: 11)

      expect(p.to_html).to eq('<p style="font-size:11pt;">Normal<em>Italic</em></p>')
    end

    it 'sets the color of colored text' do
      p = paragraph('<w:p><w:r><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>red</w:t></w:r></w:p>', font_size: 11)

      expect(p.to_html).to eq('<p style="font-size:11pt;color:#FF0000;">red</p>')
    end

    it 'sets text-align from the paragraph alignment' do
      center = paragraph('<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>x</w:t></w:r></w:p>', font_size: 11)
      plain = paragraph('<w:p><w:r><w:t>x</w:t></w:r></w:p>', font_size: 11)

      expect(center.to_html).to eq('<p style="font-size:11pt;text-align:center;">x</p>')
      expect(plain.to_html).to eq('<p style="font-size:11pt;">x</p>')
    end

    it 'uses the paragraph font size' do
      p = paragraph('<w:p><w:pPr><w:rPr><w:sz w:val="28"/></w:rPr></w:pPr><w:r><w:t>x</w:t></w:r></w:p>', font_size: 11)

      expect(p.to_html).to eq('<p style="font-size:14pt;">x</p>')
    end
  end

  describe 'styles' do
    let(:styles_configuration) do
      Docx::Elements::Containers::StylesConfiguration.new(Nokogiri::XML(<<~XML))
        <w:styles xmlns:w="#{XmlHelpers::NAMESPACES['w']}">
          <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="Heading 1"/></w:style>
          <w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/></w:style>
        </w:styles>
      XML
    end
    let(:document) { instance_double(Docx::Document, styles_configuration: styles_configuration) }
    let(:styled) { paragraph('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>', {}, document) }

    describe '#style_id' do
      it 'reads the w:pStyle value' do
        expect(styled.style_id).to eq('Heading1')
      end

      it 'is nil when the paragraph has no style' do
        expect(paragraph('<w:p><w:pPr/></w:p>').style_id).to be_nil
      end
    end

    describe '#style' do
      it 'is nil without a document' do
        expect(paragraph('<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr></w:p>').style).to be_nil
      end

      it 'resolves the style name through the document' do
        allow(document).to receive(:style_name_of).with('Heading1').and_return('Heading 1')

        expect(styled.style).to eq('Heading 1')
      end

      it 'falls back to the default paragraph style' do
        allow(document).to receive(:style_name_of).and_return(nil)
        allow(document).to receive(:default_paragraph_style).and_return('Normal')

        expect(styled.style).to eq('Normal')
      end
    end

    describe '#style=' do
      it 'sets w:pStyle from a style name' do
        styled.style = 'Heading 2'

        expect(styled.style_id).to eq('Heading2')
      end

      it 'sets w:pStyle from a style id (also as #style_id=)' do
        styled.style_id = 'Heading2'

        expect(styled.style_id).to eq('Heading2')
      end

      it 'adds w:pStyle to a paragraph without one' do
        p = paragraph('<w:p><w:pPr/></w:p>', {}, document)

        p.style = 'Heading 1'

        expect(p.node.at_xpath('w:pPr/w:pStyle/@w:val').value).to eq('Heading1')
      end

      it 'raises for an unknown style' do
        expect { styled.style = 'invalid' }.to raise_error(Docx::Errors::StyleNotFound)
      end
    end
  end
end
