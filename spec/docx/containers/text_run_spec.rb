# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Containers::TextRun do
  let(:default_formatting) { described_class::DEFAULT_FORMATTING }

  def run(xml, document_properties = {})
    described_class.new(w_node(xml), document_properties)
  end

  it 'is backed by the w:r tag' do
    expect(described_class.tag).to eq('r')
  end

  describe '#text' do
    it 'joins the w:t nodes of the run' do
      expect(run('<w:r><w:t>Hello </w:t><w:t>world</w:t></w:r>').text).to eq('Hello world')
      expect(run('<w:r><w:t>Hello</w:t></w:r>').to_s).to eq('Hello')
    end

    it 'is empty for a run without text' do
      expect(run('<w:r/>').text).to eq('')
    end

    it 'reads the text of a hyperlink' do
      expect(run('<w:hyperlink r:id="rId1"><w:r><w:t>link</w:t></w:r></w:hyperlink>').text).to eq('link')
    end
  end

  describe '#text=' do
    it 'replaces the text of a single w:t' do
      node = w_node('<w:r><w:t>old</w:t></w:r>')
      text_run = described_class.new(node)

      text_run.text = 'new'

      expect(text_run.text).to eq('new')
      expect(node.xpath('w:t').map(&:text)).to eq(['new'])
    end

    it 'adds a w:t when the run has none' do
      node = w_node('<w:r/>')
      text_run = described_class.new(node)

      text_run.text = 'new'

      expect(node.xpath('w:t').map(&:text)).to eq(['new'])
    end

    # Regression tests for #190
    it 'reflects the new text when the run had no w:t' do
      text_run = described_class.new(w_node('<w:r/>'))

      text_run.text = 'new'

      expect(text_run.text).to eq('new')
      expect(text_run.to_html).to eq('new')
    end

    it 'keeps updating the same w:t on later assignments' do
      node = w_node('<w:r/>')
      text_run = described_class.new(node)

      text_run.text = 'first'
      text_run.text = 'second'

      expect(text_run.text).to eq('second')
      expect(node.xpath('w:t').map(&:text)).to eq(['second'])
    end

    # Regression tests for #197
    context 'when the run has several w:t' do
      # The run content in document order, breaks and tabs included.
      def layout(node)
        node.xpath('.//w:t|.//w:br|.//w:tab|.//w:cr').map { |n| n.name == 't' ? n.text : "<#{n.name}>" }
      end

      it 'replaces the text from the first through the last w:t, dropping the breaks and tabs between them' do
        node = w_node('<w:r><w:rPr><w:b/></w:rPr><w:t>a</w:t><w:br/><w:t>b</w:t><w:tab/><w:t>c</w:t></w:r>')
        text_run = described_class.new(node)

        text_run.text = 'new'

        expect(text_run.text).to eq('new')
        expect(layout(node)).to eq(['new'])
        expect(text_run).to be_bolded
      end

      it 'keeps breaks outside the text, such as a trailing page break' do
        node = w_node('<w:r><w:br/><w:t>a</w:t><w:cr/><w:t>b</w:t><w:br w:type="page"/></w:r>')
        text_run = described_class.new(node)

        text_run.text = 'new'

        expect(layout(node)).to eq(['<br>', 'new', '<br>'])
        expect(node.at_xpath('w:br[@w:type="page"]')).not_to be_nil
      end

      it 'replaces the text of a hyperlink made of several runs' do
        node = w_node('<w:hyperlink r:id="rId1"><w:r><w:t>site </w:t></w:r><w:r><w:t>link</w:t></w:r></w:hyperlink>')
        text_run = described_class.new(node)

        text_run.text = 'new'

        expect(text_run.text).to eq('new')
        expect(layout(node)).to eq(['new'])
      end
    end
  end

  describe '#substitute' do
    it 'replaces matches in every w:t of the run' do
      text_run = run('<w:r><w:t>_a_ and </w:t><w:t>_a_</w:t></w:r>')

      text_run.substitute('_a_', 'x')

      expect(text_run.text).to eq('x and x')
    end

    it 'accepts a Regexp with backreferences' do
      text_run = run('<w:r><w:t>_placeholder2_</w:t></w:r>')

      text_run.substitute(/_placeholder(\d)_/, 'value\1')

      expect(text_run.text).to eq('value2')
    end
  end

  describe '#substitute_with_block' do
    it 'yields the MatchData of each match' do
      text_run = run('<w:r><w:t>line 1_placeholder2_ line 2_placeholder3_</w:t></w:r>')

      text_run.substitute_with_block(/_placeholder(\d)_/) { |match| "_replacement_#{match[1]}" }

      expect(text_run.text).to eq('line 1_replacement_2 line 2_replacement_3')
    end
  end

  describe '#formatting' do
    it 'is the default formatting for a plain run' do
      text_run = run('<w:r><w:t>x</w:t></w:r>')

      expect(text_run.formatting).to eq(default_formatting)
      expect(text_run).not_to be_italicized
      expect(text_run).not_to be_bolded
      expect(text_run).not_to be_underlined
      expect(text_run).not_to be_striked
    end

    it 'detects italic' do
      text_run = run('<w:r><w:rPr><w:i/></w:rPr><w:t>x</w:t></w:r>')

      expect(text_run.formatting).to eq(default_formatting.merge(italic: true))
      expect(text_run).to be_italicized
    end

    it 'detects bold' do
      text_run = run('<w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r>')

      expect(text_run.formatting).to eq(default_formatting.merge(bold: true))
      expect(text_run).to be_bolded
    end

    it 'detects underline' do
      text_run = run('<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>x</w:t></w:r>')

      expect(text_run.formatting).to eq(default_formatting.merge(underline: true))
      expect(text_run).to be_underlined
    end

    it 'detects strike' do
      text_run = run('<w:r><w:rPr><w:strike/></w:rPr><w:t>x</w:t></w:r>')

      expect(text_run.formatting).to eq(default_formatting.merge(strike: true))
      expect(text_run).to be_striked
    end

    it 'detects combined formatting' do
      text_run = run('<w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr><w:t>x</w:t></w:r>')

      expect(text_run.formatting).to eq(default_formatting.merge(italic: true, bold: true, underline: true))
    end
  end

  # ECMA-376 17.3.2.38: w:sz is in half-points, so 14pt is stored as 28.
  describe '#font_size' do
    it 'reads the run size in points' do
      expect(run('<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>x</w:t></w:r>', font_size: 11).font_size).to eq(16)
    end

    it 'falls back to the document font size' do
      expect(run('<w:r><w:t>x</w:t></w:r>', font_size: 11).font_size).to eq(11)
    end

    it 'is nil without a run size or document font size' do
      expect(run('<w:r><w:t>x</w:t></w:r>').font_size).to be_nil
    end
  end

  describe 'hyperlinks' do
    let(:properties) { { font_size: 11, hyperlinks: { 'rId1' => 'http://www.google.com/' } } }
    let(:external) { run('<w:hyperlink r:id="rId1"><w:r><w:t>link</w:t></w:r></w:hyperlink>', properties) }
    let(:internal) { run('<w:hyperlink w:anchor="_Toc1"><w:r><w:t>link</w:t></w:r></w:hyperlink>', properties) }

    it 'treats a hyperlink with a relationship id as an external hyperlink' do
      expect(external).to be_hyperlink
      expect(external).to be_external_link
      expect(external.hyperlink_id).to eq('rId1')
      expect(external.href).to eq('http://www.google.com/')
    end

    it 'does not treat an anchor hyperlink as a hyperlink' do
      expect(internal).not_to be_hyperlink
      expect(internal).not_to be_external_link
    end

    it 'does not treat a plain run as a hyperlink' do
      expect(run('<w:r><w:t>x</w:t></w:r>')).not_to be_hyperlink
    end
  end

  describe '#to_html' do
    let(:properties) { { font_size: 11, hyperlinks: { 'rId1' => 'http://www.google.com/' } } }

    it 'returns plain text for an unformatted run' do
      expect(run('<w:r><w:t>plain</w:t></w:r>', properties).to_html).to eq('plain')
    end

    it 'wraps italic text in em' do
      expect(run('<w:r><w:rPr><w:i/></w:rPr><w:t>Italic</w:t></w:r>', properties).to_html).to eq('<em>Italic</em>')
    end

    it 'wraps bold text in strong' do
      expect(run('<w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r>', properties).to_html).to eq('<strong>Bold</strong>')
    end

    it 'wraps struck text in s' do
      expect(run('<w:r><w:rPr><w:strike/></w:rPr><w:t>Strike</w:t></w:r>', properties).to_html).to eq('<s>Strike</s>')
    end

    it 'underlines with a styled span' do
      html = run('<w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>Underline</w:t></w:r>', properties).to_html

      expect(html).to eq('<span style="text-decoration:underline;">Underline</span>')
    end

    it 'nests combined formatting as span > strong > em' do
      html = run('<w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr><w:t>all</w:t></w:r>', properties).to_html

      expect(html).to eq('<span style="text-decoration:underline;"><strong><em>all</em></strong></span>')
    end

    it 'sets the font size only when it differs from the document font size' do
      larger = run('<w:r><w:rPr><w:sz w:val="32"/></w:rPr><w:t>big</w:t></w:r>', properties)
      same = run('<w:r><w:rPr><w:sz w:val="22"/></w:rPr><w:t>same</w:t></w:r>', properties)

      expect(larger.to_html).to eq('<span style="font-size:16pt;">big</span>')
      expect(same.to_html).to eq('same')
    end

    it 'renders an external hyperlink as an anchor opening in a new tab' do
      html = run('<w:hyperlink r:id="rId1"><w:r><w:t>hyperlink</w:t></w:r></w:hyperlink>', properties).to_html

      expect(html).to eq('<a href="http://www.google.com/" target="_blank">hyperlink</a>')
    end
  end
end
