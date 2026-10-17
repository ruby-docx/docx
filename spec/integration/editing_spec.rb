# frozen_string_literal: true

require 'spec_helper'
require 'docx'

# End-to-end editing scenarios against real .docx files, including saving and
# reopening. Unit-level behavior lives in spec/docx.
describe 'Editing a document' do
  describe 'editing.docx' do
    let(:doc) { Docx::Document.open(fixture_path('editing.docx')) }

    it 'copies and inserts a paragraph' do
      first = doc.paragraphs.first

      first.copy.insert_after(first)

      expect(doc.paragraphs.map(&:text).first(2)).to eq(['test text', 'test text'])
      expect(doc.paragraphs.size).to eq(4)
    end

    it 'changes, blanks and removes paragraphs' do
      doc.paragraphs.first.text = 'the real test'
      expect(doc.paragraphs.first.text).to eq('the real test')

      doc.paragraphs.first.blank!
      expect(doc.paragraphs.first.text).to eq('')

      expect { doc.paragraphs.first.remove! }.to change { doc.paragraphs.size }.by(-1)
    end

    it 'inserts text around bookmarks' do
      doc.bookmarks['beginning_bookmark'].insert_text_before('foo')
      doc.bookmarks['end_bookmark'].insert_text_after('bar')

      expect(doc.paragraphs.first.text).to eq('footest textbar')
    end

    it 'inserts multiple lines at an isolated bookmark' do
      lines = ['replacement test', 'second paragraph test', 'and a third paragraph test']

      doc.bookmarks['isolated_bookmark'].insert_multiple_lines(lines)

      expect(doc.paragraphs.map(&:text)[2..]).to eq(lines)
    end

    it 'inserts multiple lines at a bookmark splitting a word' do
      lines = ['replacement test', 'second paragraph test', 'and a third paragraph test']

      doc.bookmarks['word_splitting_bookmark'].insert_multiple_lines(lines)

      expect(doc.paragraphs.map(&:text)[1, 3]).to eq(lines)
    end
  end

  describe 'substitution.docx' do
    let(:doc) { Docx::Document.open(fixture_path('substitution.docx')) }
    let(:paragraph) { doc.paragraphs[1] }

    it 'substitutes placeholders run by run, preserving formatting' do
      paragraph.each_text_run do |run|
        run.substitute('_placeholder2_', 'same paragraph')
        run.substitute('_placeholder3_', 'yet the same paragraph')
      end

      expect(paragraph.text).to eq('Multi-line paragraph line 1same paragraph line 2yet the same paragraph line3 ')
    end

    it 'substitutes placeholders with a block' do
      paragraph.each_text_run do |run|
        run.substitute_with_block(/_placeholder(\d)_/) { |match| "_replacement_#{match[1]}" }
      end

      expect(paragraph.text).to eq('Multi-line paragraph line 1_replacement_2 line 2_replacement_3 line3 ')
    end
  end

  # Regression test for #147
  describe 'split_placeholder.docx' do
    it 'replaces a placeholder Word split across runs and persists it' do
      doc = Docx::Document.open(fixture_path('split_placeholder.docx'))
      paragraph = doc.paragraphs.find { |p| p.text.include?('{{') }
      expect(paragraph.text_runs.map(&:text)).to eq(['Hello ', '{{fi', 'rst_na', 'me}}', '!'])

      paragraph.substitute('{{first_name}}', 'World')
      doc.save(tmp_path('saved.docx'))

      reopened = Docx::Document.open(tmp_path('saved.docx'))
      expect(reopened.paragraphs.map(&:text)).to include('Hello World!')
    end
  end

  # Regression test for #111: Google Docs places bookmarks directly under w:body.
  describe 'bookmark_outside_paragraph.docx' do
    let(:doc) { Docx::Document.open(fixture_path('bookmark_outside_paragraph.docx')) }

    it 'inserts text at a block-level bookmark and persists it' do
      doc.bookmarks['bookmark_1'].insert_text_after('INSERTED ')
      doc.save(tmp_path('saved.docx'))

      expect(Docx::Document.open(tmp_path('saved.docx')).paragraphs.map(&:text)).to eq(['INSERTED bookmark_1'])
    end

    it 'inserts multiple lines without destroying the following paragraph' do
      doc.bookmarks['bookmark_1'].insert_multiple_lines(%w[line1 line2 line3])

      expect(doc.paragraphs.map(&:text)).to eq(%w[line1 line2 line3 bookmark_1])
    end
  end

  describe 'multi_doc_bookmarks.docx' do
    it 'edits bookmarks in headers and footers and persists them' do
      doc = Docx::Document.open(fixture_path('multi_doc_bookmarks.docx'))

      doc.bookmarks['header_bookmark'].insert_text_after('Inserted ')
      doc.bookmarks['footer_bookmark'].insert_text_after('Inserted ')
      expect(doc.headers['header1'].text).to eq('Inserted Hello from the header.')
      expect(doc.footers['footer1'].text).to eq('Inserted Hello from the footer.')

      doc.save(tmp_path('saved.docx'))

      reopened = Docx::Document.open(tmp_path('saved.docx'))
      expect(reopened.headers['header1'].text).to eq('Inserted Hello from the header.')
      expect(reopened.footers['footer1'].text).to eq('Inserted Hello from the footer.')
    end
  end
end
