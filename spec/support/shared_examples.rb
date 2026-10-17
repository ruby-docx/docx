# frozen_string_literal: true

# Expects +doc+ to be spec/fixtures/basic.docx, however it was opened.
RSpec.shared_examples_for 'reading basic.docx' do
  it 'reads the paragraphs' do
    expect(doc.paragraphs.size).to eq(2)
    expect(doc.paragraphs.first.text).to eq('hello')
    expect(doc.paragraphs.last.text).to eq('world')
    expect(doc.text).to eq("hello\nworld")
  end

  it 'reads the bookmarks' do
    expect(doc.bookmarks.keys).to eq(['test_bookmark'])
  end

  it 'reads paragraphs made of unformatted text runs' do
    doc.each_paragraph do |paragraph|
      expect(paragraph).to be_an_instance_of(Docx::Elements::Containers::Paragraph)
      paragraph.each_text_run do |run|
        expect(run).to be_an_instance_of(Docx::Elements::Containers::TextRun)
        expect(run.formatting).to eq(Docx::Elements::Containers::TextRun::DEFAULT_FORMATTING)
      end
    end
  end
end

# Expects +doc+ to be an opened Docx::Document.
RSpec.shared_examples_for 'saving to file' do
  it 'saves to a file path that can be reopened' do
    path = tmp_path('saved.docx')
    doc.save(path)

    expect(Docx::Document.open(path).paragraphs.size).to eq(doc.paragraphs.size)
  end
end
