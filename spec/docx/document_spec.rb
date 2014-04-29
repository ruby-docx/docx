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
      @doc.paragraphs.size.should eq 2
      @doc.paragraphs.first.text.should eq 'hello'
      @doc.paragraphs.last.text.should eq 'world'
      @doc.text.should eq "hello\nworld"
    end

    it 'should read bookmarks' do
      @doc.bookmarks.size.should eq 1
      @doc.bookmarks['test_bookmark'].should_not be_nil
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
          tr.formatting.should eq Docx::Elements::Containers::TextRun::DEFAULT_FORMATTING
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
      @doc.tables[0].rows[0].cells[0].text.should eq "ENGLISH"
      @doc.tables[0].rows[0].cells[1].text.should eq "FRANÇAIS"
      @doc.tables[1].rows[0].cells[0].text.should eq "Second table"
      @doc.tables[1].rows[0].cells[1].text.should eq "Second tableau"
      @doc.tables[0].columns[0].cells[5].text.should eq "aphids"
      @doc.tables[0].columns[1].cells[5].text.should eq "puceron"
    end

    it "should read embedded links" do
      @doc.tables[0].columns[1].cells[1].text.should =~ /^Directive/
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
      new_p.should_not be_nil
      new_p.should_not eq(old_p)
    end

    it 'allows insertion of text' do
      @doc.paragraphs.size.should eq 3
      first_p = @doc.paragraphs.first
      new_p = first_p.copy
      new_p.insert_after first_p
      @doc.paragraphs.size.should eq 4
    end

    it 'should change text' do
      @doc.paragraphs.first.text.should eq 'test text'
      @doc.paragraphs.first.text = 'the real test'
      @doc.paragraphs.first.text.should eq 'the real test'
    end

    it 'should allow insertion of text before a bookmark' do
      @doc.paragraphs.first.text.should eq 'test text'
      @doc.bookmarks['beginning_bookmark'].insert_text_before('foo')
      @doc.paragraphs.first.text.should eq 'footest text'
    end

    it 'should allow insertion of text after a bookmark' do
      @doc.paragraphs.first.text.should eq 'test text'
      @doc.bookmarks['end_bookmark'].insert_text_after('bar')
      @doc.paragraphs.first.text.should eq 'test textbar'
    end

    it 'should allow multiple lines of text to be inserted at a bookmark' do
      @doc.paragraphs.last.text.should eq ''
      new_lines = ['replacement test', 'second paragraph test', 'and a third paragraph test']
      @doc.bookmarks['isolated_bookmark'].insert_multiple_lines(new_lines)
      new_lines.each_index do |line|
        @doc.paragraphs[line + 2].text.should eq new_lines[line]
      end
    end

    it 'should allow multi-line insertion with replacement' do
      @doc.paragraphs[1].text.should eq 'placeholder text'
      new_lines = ['replacement test', 'second paragraph test', 'and a third paragraph test']
      @doc.bookmarks['word_splitting_bookmark'].insert_multiple_lines(new_lines)
      new_lines.each_index do |line|
        @doc.paragraphs[line + 1].text.should eq new_lines[line]
      end
    end

    it 'should allow content deletion' do
      @doc.paragraphs.first.text.should eq 'test text'
      @doc.paragraphs.first.blank!
      @doc.paragraphs.first.text.should eq ''
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
      @doc.paragraphs.size.should eq 11
      @doc.paragraphs[0].text.should eq 'Normal'
      @doc.paragraphs[1].text.should eq 'Italic'
      @doc.paragraphs[2].text.should eq 'Bold'
      @doc.paragraphs[3].text.should eq 'Underline'
      @doc.paragraphs[4].text.should eq 'Normal'
      @doc.paragraphs[5].text.should eq 'This is a sentence with all formatting options in the middle of the sentence.'
      @doc.paragraphs[6].text.should eq 'This is a centered paragraph.'
      @doc.paragraphs[7].text.should eq 'This paragraph is aligned left.'
      @doc.paragraphs[8].text.should eq 'This paragraph is aligned right.'
      @doc.paragraphs[9].text.should eq 'This paragraph is 14 points.'
      @doc.paragraphs[10].text.should eq 'This paragraph has a word at 16 points.'
    end

    it 'should contain a paragraph with multiple text runs' do

    end

    it 'should detect normal formatting' do
      [0, 4].each do |i|
        @formatting[i][0].should eq @default_formatting
        @doc.paragraphs[i].text_runs[0].italicized?.should be_false
        @doc.paragraphs[i].text_runs[0].bolded?.should be_false
        @doc.paragraphs[i].text_runs[0].underlined?.should be_false
      end
    end

    it 'should detect italic formatting' do
      @formatting[1][0].should eq @only_italic
      @doc.paragraphs[1].text_runs[0].italicized?.should be_true
      @doc.paragraphs[1].text_runs[0].bolded?.should be_false
      @doc.paragraphs[1].text_runs[0].underlined?.should be_false
    end

    it 'should detect bold formatting' do
      @formatting[2][0].should eq @only_bold
      @doc.paragraphs[2].text_runs[0].italicized?.should be_false
      @doc.paragraphs[2].text_runs[0].bolded?.should be_true
      @doc.paragraphs[2].text_runs[0].underlined?.should be_false
    end

    it 'should detect underline formatting' do
      @formatting[3][0].should eq @only_underline
      @doc.paragraphs[3].text_runs[0].italicized?.should be_false
      @doc.paragraphs[3].text_runs[0].bolded?.should be_false
      @doc.paragraphs[3].text_runs[0].underlined?.should be_true
    end

    it 'should detect mixed formatting' do
      @formatting[5][0].should eq @default_formatting
      @doc.paragraphs[5].text_runs[0].italicized?.should be_false
      @doc.paragraphs[5].text_runs[0].bolded?.should be_false
      @doc.paragraphs[5].text_runs[0].underlined?.should be_false
      
      @formatting[5][1].should eq @all_formatted
      @doc.paragraphs[5].text_runs[1].italicized?.should be_true
      @doc.paragraphs[5].text_runs[1].bolded?.should be_true
      @doc.paragraphs[5].text_runs[1].underlined?.should be_true
      
      @formatting[5][2].should eq @default_formatting
      @doc.paragraphs[5].text_runs[2].italicized?.should be_false
      @doc.paragraphs[5].text_runs[2].bolded?.should be_false
      @doc.paragraphs[5].text_runs[2].underlined?.should be_false
    end

    it 'should detect centered paragraphs' do
      @doc.paragraphs[5].aligned_center?.should be_false
      @doc.paragraphs[6].aligned_center?.should be_true
      @doc.paragraphs[7].aligned_center?.should be_false
    end

    it 'should detect left justified paragraphs' do
      @doc.paragraphs[6].aligned_left?.should be_false
      @doc.paragraphs[7].aligned_left?.should be_true
      @doc.paragraphs[8].aligned_left?.should be_false
    end

    it 'should detect right justified paragraphs' do
      @doc.paragraphs[7].aligned_right?.should be_false
      @doc.paragraphs[8].aligned_right?.should be_true
      @doc.paragraphs[9].aligned_right?.should be_false
    end

    # ECMA-376 Office Open XML spec (4th edition), 17.3.2.38, size is
    # defined in half-points, meaning 14pt text returns a value of 28.
    # http://www.ecma-international.org/publications/standards/Ecma-376.htm
    it 'should return proper font size for paragraphs' do
      @doc.font_size.should eq 11
      @doc.paragraphs[5].font_size.should eq 11
      paragraph = @doc.paragraphs[9]
      paragraph.font_size.should eq 14
      paragraph.text_runs[0].font_size.should eq 14
    end

    it 'should return proper font size for runs' do
      @doc.font_size.should eq 11
      paragraph = @doc.paragraphs[10]
      paragraph.font_size.should eq 11
      text_runs = paragraph.text_runs
      text_runs[0].font_size.should eq 11
      text_runs[1].font_size.should eq 16
      text_runs[2].font_size.should eq 11
      text_runs[3].font_size.should eq 11
      text_runs[4].font_size.should eq 11
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
      @new_doc.paragraphs.size.should eq @doc.paragraphs.size
    end

    it 'should save to a tempfile' do
      temp_file = Tempfile.new(['docx_gem', '.docx'])
      @new_doc_path = temp_file.path
      @doc.save(@new_doc_path)
      @new_doc = Docx::Document.open(@new_doc_path)
      @new_doc.paragraphs.size.should eq @doc.paragraphs.size

      temp_file.close
      temp_file.unlink
      # ensure temp file has been removed
      File.exists?(@new_doc_path).should be_false
    end

    after do
      if File.exists?(@new_doc_path)
        File.delete(@new_doc_path)
      end
    end
  end
end