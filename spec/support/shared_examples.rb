# frozen_string_literal: true

RSpec.shared_examples_for 'reading' do
  it 'should read the document' do
    expect(@doc.paragraphs.size).to eq(2)
    expect(@doc.paragraphs.first.text).to eq('hello')
    expect(@doc.paragraphs.last.text).to eq('world')
    expect(@doc.text).to eq("hello\nworld")
  end

  it 'should read bookmarks' do
    expect(@doc.bookmarks.size).to eq(1)
    expect(@doc.bookmarks['test_bookmark']).to_not eq(nil)
  end

  it 'should have paragraphs' do
    @doc.each_paragraph do |p|
      expect(p).to be_an_instance_of(Docx::Elements::Containers::Paragraph)
    end
  end

  it 'should have properly formatted text runs' do
    @doc.each_paragraph do |p|
      p.each_text_run do |tr|
        expect(tr).to be_an_instance_of(Docx::Elements::Containers::TextRun)
        expect(tr.formatting).to eq(Docx::Elements::Containers::TextRun::DEFAULT_FORMATTING)
      end
    end
  end

  describe '#font_size' do
    context 'When a docx files has no styles.xml' do
      before do
        @doc = Docx::Document.new(@fixtures_path + '/no_styles.docx')
      end

      it 'should raise an error' do
        expect(@doc.font_size).to be_nil
      end
    end
  end
end

RSpec.shared_examples_for 'saving to file' do
  it 'should save to a normal file path' do
    @new_doc_path = @fixtures_path + '/new_save.docx'
    @doc.save(@new_doc_path)
    @new_doc = Docx::Document.open(@new_doc_path)
    expect(@new_doc.paragraphs.size).to eq(@doc.paragraphs.size)
  end

  it 'should save to a tempfile' do
    temp_file = Tempfile.new(['docx_gem', '.docx'])
    @new_doc_path = temp_file.path
    @doc.save(@new_doc_path)
    @new_doc = Docx::Document.open(@new_doc_path)
    expect(@new_doc.paragraphs.size).to eq(@doc.paragraphs.size)

    temp_file.close
    temp_file.unlink
    # ensure temp file has been removed
    expect(File.exist?(@new_doc_path)).to eq(false)
  end

  after do
    File.delete(@new_doc_path) if File.exist?(@new_doc_path)
  end
end
