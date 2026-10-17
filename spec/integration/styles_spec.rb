# frozen_string_literal: true

require 'spec_helper'
require 'docx'

# Reading and changing styles in a real Word document.
describe 'Styles in styles.docx' do
  let(:doc) { Docx::Document.open(fixture_path('styles.docx')) }

  it 'reads the style of every paragraph' do
    expect(doc.paragraphs.map(&:style)).to eq([
      'Title', 'Subtitle', 'Author', 'Date', 'Compact',
      'Heading 1', 'Heading 2', 'Heading 3', 'Heading 4', 'Heading 5',
      'Heading 6', 'Heading 7', 'Heading 8', 'Heading 9',
      'First Paragraph', 'Body Text', 'Block Text', 'Table Caption', 'Image Caption',
      'Definition Term', 'Definition', 'Definition Term', 'Definition'
    ])
    expect(doc.paragraphs.map(&:style_id)).to eq(%w[
      Title Subtitle Author Date Compact
      Heading1 Heading2 Heading3 Heading4 Heading5
      Heading6 Heading7 Heading8 Heading9
      FirstParagraph BodyText BlockText TableCaption ImageCaption
      DefinitionTerm Definition DefinitionTerm Definition
    ])
  end

  it 'sets paragraph styles by name and by id' do
    doc.paragraphs.each { |p| p.style = 'Heading 1' }
    expect(doc.paragraphs.map(&:style).uniq).to eq(['Heading 1'])

    doc.paragraphs.each { |p| p.style_id = 'Heading2' }
    expect(doc.paragraphs.map(&:style).uniq).to eq(['Heading 2'])
  end

  it 'adds a style, applies it, and removes it again across saves' do
    doc.styles_configuration.add_style('Red', name: 'Red', font_color: 'FF0000', font_size: 20)
    doc.paragraphs[5].style = 'Red'
    doc.save(tmp_path('with_red.docx'))

    with_red = Docx::Document.open(tmp_path('with_red.docx'))
    expect(with_red.paragraphs[5].style).to eq('Red')
    with_red.styles_configuration.remove_style('Red')
    with_red.save(tmp_path('without_red.docx'))

    without_red = Docx::Document.open(tmp_path('without_red.docx'))
    expect { without_red.styles_configuration.style_of('Red') }.to raise_error(Docx::Errors::StyleNotFound)
  end
end
