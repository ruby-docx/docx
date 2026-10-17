# frozen_string_literal: true

require 'spec_helper'
require 'docx'

# Reading formatting and rendering html from a real Word document.
describe 'Reading formatting.docx' do
  let(:doc) { Docx::Document.open(fixture_path('formatting.docx')) }
  let(:paragraphs) { doc.paragraphs }
  let(:default_formatting) { Docx::Elements::Containers::TextRun::DEFAULT_FORMATTING }

  it 'reads the text of every paragraph' do
    expect(paragraphs.map(&:text)).to eq([
      'Normal',
      'Italic',
      'Bold',
      'Underline',
      'Normal',
      'This is a sentence with all formatting options in the middle of the sentence.',
      'This is a centered paragraph.',
      'This paragraph is aligned left.',
      'This paragraph is aligned right.',
      'This paragraph is 14 points.',
      'This paragraph has a word at 16 points.',
      'This sentence has different formatting in different places.',
      'This sentence has a hyperlink.',
      'Strike',
      'Red color'
    ])
  end

  it 'reads run formatting' do
    expect(paragraphs[0].text_runs[0].formatting).to eq(default_formatting)
    expect(paragraphs[1].text_runs[0].formatting).to eq(default_formatting.merge(italic: true))
    expect(paragraphs[2].text_runs[0].formatting).to eq(default_formatting.merge(bold: true))
    expect(paragraphs[3].text_runs[0].formatting).to eq(default_formatting.merge(underline: true))
    expect(paragraphs[5].text_runs.first(3).map(&:formatting)).to eq([
      default_formatting,
      default_formatting.merge(italic: true, bold: true, underline: true),
      default_formatting
    ])
  end

  it 'reads paragraph alignment' do
    expect(paragraphs[6..8].map { |p| [p.aligned_left?, p.aligned_center?, p.aligned_right?] }).to eq([
      [false, true, false],
      [true, false, false],
      [false, false, true]
    ])
  end

  it 'reads font sizes' do
    expect(doc.font_size).to eq(11)
    expect(paragraphs[9].font_size).to eq(14)
    expect(paragraphs[9].text_runs[0].font_size).to eq(14)
    expect(paragraphs[10].font_size).to eq(11)
    expect(paragraphs[10].text_runs.map(&:font_size)).to eq([11, 16, 11, 11, 11])
  end

  describe 'html output' do
    it 'renders the opening paragraphs' do
      expect(doc.to_html).to start_with(
        %(<p style="font-size:11pt;">Normal</p>\n<p style="font-size:11pt;"><em>Italic</em></p>\n<p style="font-size:11pt;"><strong>Bold</strong></p>)
      )
    end

    it 'renders one p tag per paragraph' do
      expect(doc.to_html.scan('<p').size).to eq(paragraphs.size)
    end

    it 'renders mixed formatting within a sentence' do
      html = paragraphs[11].to_html

      expect(html).to include('<em>sentence</em>', '<strong>formatting</strong>',
                              '<span style="text-decoration:underline;">different</span>')
    end

    it 'renders all formatting options on a single run' do
      expect(paragraphs[5].to_html).to include('<span style="text-decoration:underline;"><strong><em>all</em></strong></span>')
    end

    it 'renders alignment, size, color and strike' do
      expect(paragraphs[6].to_html).to include('text-align:center')
      expect(paragraphs[8].to_html).to include('text-align:right')
      expect(paragraphs[9].to_html).to include('font-size:14pt')
      expect(paragraphs[10].to_html).to include('<span style="font-size:16pt;">')
      expect(paragraphs[13].to_html).to include('<s>Strike</s>')
      expect(paragraphs[14].to_html).to start_with('<p style="font-size:11pt;color:#FF0000;">')
    end

    it 'renders hyperlinks as anchors' do
      expect(doc.to_html).to include('<a href="http://www.google.com/" target="_blank">hyperlink</a>')
    end
  end
end
