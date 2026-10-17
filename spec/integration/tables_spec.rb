# frozen_string_literal: true

require 'spec_helper'
require 'docx'

# Reading tables from a real Word document.
describe 'Reading tables.docx' do
  let(:doc) { Docx::Document.open(fixture_path('tables.docx')) }
  let(:first) { doc.tables[0] }
  let(:second) { doc.tables[1] }

  it 'reads the table dimensions' do
    expect([first.row_count, first.column_count]).to eq([171, 2])
    expect([second.row_count, second.column_count]).to eq([2, 2])
  end

  it 'reads cells by row and by column' do
    expect(first.rows[0].cells.map(&:text)).to eq(%w[ENGLISH FRANÇAIS])
    expect(second.rows[0].cells.map(&:text)).to eq(['Second table', 'Second tableau'])
    expect(first.columns[0].cells[5].text).to eq('aphids')
    expect(first.columns[1].cells[5].text).to eq('puceron')
  end

  it 'reads text inside embedded links' do
    expect(first.columns[1].cells[1].text).to match(/^Directive/)
  end

  # Regression test for #160
  it 'renders html for cells containing hyperlinks' do
    html = first.rows[1].cells[1].paragraphs.map(&:to_html).join

    expect(html).to include('<a href="http://google.com/"')
  end
end
