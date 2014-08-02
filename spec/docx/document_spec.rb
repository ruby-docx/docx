# coding: utf-8
require 'docx'
require 'tempfile'

describe Docx::Document do
  before(:all) do
    @fixtures_path = "spec/fixtures"
  end

  describe 'reading' do
    before do
      @doc = Docx::Document.open(@fixtures_path + '/basic.docx')
    end

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
  end

  describe 'read tables' do
    before do
      @doc = Docx::Document.open(@fixtures_path + '/tables.docx')
    end

    it "should have tables with rows and cells" do
      expect(@doc.tables.count).to eq 2
      @doc.tables.each do |table|
        expect(table).to be_an_instance_of(Docx::Elements::Containers::Table)
        table.rows.each do |row|
          expect(row).to be_an_instance_of(Docx::Elements::Containers::TableRow)
          row.cells.each do |cell|
            expect(cell).to be_an_instance_of(Docx::Elements::Containers::TableCell)
          end
        end
      end
    end

    it "should have tables with columns and cells" do
      @doc.tables.each do |table|
        table.columns.each do |column|
          expect(column).to be_an_instance_of(Docx::Elements::Containers::TableColumn)
          column.cells.each do |cell|
            expect(cell).to be_an_instance_of(Docx::Elements::Containers::TableCell)
          end
        end
      end
    end

    it "should have proper count" do
      expect(@doc.tables[0].row_count).to eq 171
      expect(@doc.tables[1].row_count).to eq 2
      expect(@doc.tables[0].column_count).to eq 2
      expect(@doc.tables[1].column_count).to eq 2
    end

    it "should have tables with proper text" do
      expect(@doc.tables[0].rows[0].cells[0].text).to eq "ENGLISH"
      expect(@doc.tables[0].rows[0].cells[1].text).to eq "FRANÇAIS"
      expect(@doc.tables[1].rows[0].cells[0].text).to eq "Second table"
      expect(@doc.tables[1].rows[0].cells[1].text).to eq "Second tableau"
      expect(@doc.tables[0].columns[0].cells[5].text).to eq "aphids"
      expect(@doc.tables[0].columns[1].cells[5].text).to eq "puceron"
    end

    it "should read embedded links" do
      expect(@doc.tables[0].columns[1].cells[1].text).to match(/^Directive/)
    end
  end

  describe 'editing'  do
    before do
      @doc = Docx::Document.open(@fixtures_path + '/editing.docx')
    end

    it 'should copy paragraphs' do
      old_p = @doc.paragraphs.first
      new_p = old_p.copy
      expect(new_p).to be_an_instance_of(Docx::Elements::Containers::Paragraph)
      expect(new_p).not_to eq(nil)
      expect(new_p).not_to eq(old_p)
    end

    it 'allows insertion of text' do
      expect(@doc.paragraphs.size).to eq(3)
      first_p = @doc.paragraphs.first
      new_p = first_p.copy
      new_p.insert_after first_p
      expect(@doc.paragraphs.size).to eq(4)
    end

    it 'should change text' do
      expect(@doc.paragraphs.first.text).to eq('test text')
      @doc.paragraphs.first.text = 'the real test'
      expect(@doc.paragraphs.first.text).to eq('the real test')
    end

    it 'should allow insertion of text before a bookmark' do
      expect(@doc.paragraphs.first.text).to eq('test text')
      @doc.bookmarks['beginning_bookmark'].insert_text_before('foo')
      expect(@doc.paragraphs.first.text).to eq('footest text')
    end

    it 'should allow insertion of text after a bookmark' do
      expect(@doc.paragraphs.first.text).to eq('test text')
      @doc.bookmarks['end_bookmark'].insert_text_after('bar')
      expect(@doc.paragraphs.first.text).to eq('test textbar')
    end

    it 'should allow multiple lines of text to be inserted at a bookmark' do
      expect(@doc.paragraphs.last.text).to eq('')
      new_lines = ['replacement test', 'second paragraph test', 'and a third paragraph test']
      @doc.bookmarks['isolated_bookmark'].insert_multiple_lines(new_lines)
      new_lines.each_index do |line|
        expect(@doc.paragraphs[line + 2].text).to eq(new_lines[line])
      end
    end

    it 'should allow multi-line insertion with replacement' do
      expect(@doc.paragraphs[1].text).to eq('placeholder text')
      new_lines = ['replacement test', 'second paragraph test', 'and a third paragraph test']
      @doc.bookmarks['word_splitting_bookmark'].insert_multiple_lines(new_lines)
      new_lines.each_index do |line|
        expect(@doc.paragraphs[line + 1].text).to eq(new_lines[line])
      end
    end

    it 'should allow content deletion' do
      expect(@doc.paragraphs.first.text).to eq('test text')
      @doc.paragraphs.first.blank!
      expect(@doc.paragraphs.first.text).to eq('')
    end
  end

  describe 'read formatting' do
    before do
      @doc = Docx::Document.open(@fixtures_path + '/formatting.docx')
      @formatting = @doc.paragraphs.map { |p| p.text_runs.map(&:formatting) }
      @default_formatting = Docx::Elements::Containers::TextRun::DEFAULT_FORMATTING
      @only_italic = @default_formatting.merge italic: true
      @only_bold = @default_formatting.merge bold: true
      @only_underline = @default_formatting.merge underline: true
      @all_formatted = @default_formatting.merge italic: true, bold: true, underline: true
    end

    it 'should have the correct text' do
      expect(@doc.paragraphs.size).to eq(6)
      expect(@doc.paragraphs[0].text).to eq('Normal')
      expect(@doc.paragraphs[1].text).to eq('Italic')
      expect(@doc.paragraphs[2].text).to eq('Bold')
      expect(@doc.paragraphs[3].text).to eq('Underline')
      expect(@doc.paragraphs[4].text).to eq('Normal')
      expect(@doc.paragraphs[5].text).to eq('This is a sentence with all formatting options in the middle of the sentence.')
    end

    it 'should contain a paragraph with multiple text runs' do

    end

    it 'should detect normal formatting' do
      [0, 4].each do |i|
        expect(@formatting[i][0]).to eq(@default_formatting)
        expect(@doc.paragraphs[i].text_runs[0].italicized?).to eq(false)
        expect(@doc.paragraphs[i].text_runs[0].bolded?).to eq(false)
        expect(@doc.paragraphs[i].text_runs[0].underlined?).to eq(false)
      end
    end

    it 'should detect italic formatting' do
      expect(@formatting[1][0]).to eq(@only_italic)
      expect(@doc.paragraphs[1].text_runs[0].italicized?).to eq(true)
      expect(@doc.paragraphs[1].text_runs[0].bolded?).to eq(false)
      expect(@doc.paragraphs[1].text_runs[0].underlined?).to eq(false)
    end

    it 'should detect bold formatting' do
      expect(@formatting[2][0]).to eq(@only_bold)
      expect(@doc.paragraphs[2].text_runs[0].italicized?).to eq(false)
      expect(@doc.paragraphs[2].text_runs[0].bolded?).to eq(true)
      expect(@doc.paragraphs[2].text_runs[0].underlined?).to eq(false)
    end

    it 'should detect underline formatting' do
      expect(@formatting[3][0]).to eq(@only_underline)
      expect(@doc.paragraphs[3].text_runs[0].italicized?).to eq(false)
      expect(@doc.paragraphs[3].text_runs[0].bolded?).to eq(false)
      expect(@doc.paragraphs[3].text_runs[0].underlined?).to eq(true)
    end

    it 'should detect mixed formatting' do
      expect(@formatting[5][0]).to eq(@default_formatting)
      expect(@doc.paragraphs[5].text_runs[0].italicized?).to eq(false)
      expect(@doc.paragraphs[5].text_runs[0].bolded?).to eq(false)
      expect(@doc.paragraphs[5].text_runs[0].underlined?).to eq(false)
      
      expect(@formatting[5][1]).to eq(@all_formatted)
      expect(@doc.paragraphs[5].text_runs[1].italicized?).to eq(true)
      expect(@doc.paragraphs[5].text_runs[1].bolded?).to eq(true)
      expect(@doc.paragraphs[5].text_runs[1].underlined?).to eq(true)
      
      expect(@formatting[5][2]).to eq(@default_formatting)
      expect(@doc.paragraphs[5].text_runs[2].italicized?).to eq(false)
      expect(@doc.paragraphs[5].text_runs[2].bolded?).to eq(false)
      expect(@doc.paragraphs[5].text_runs[2].underlined?).to eq(false)
    end
  end

  describe 'saving' do
    before do
      @doc = Docx::Document.open(@fixtures_path + '/saving.docx')
    end

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
      expect(File.exists?(@new_doc_path)).to eq(false)
    end

    after do
      if File.exists?(@new_doc_path)
        File.delete(@new_doc_path)
      end
    end
  end
end
