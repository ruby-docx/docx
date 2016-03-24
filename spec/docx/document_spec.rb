# coding: utf-8
require 'spec_helper'
require 'docx'
require 'tempfile'

RSpec.describe Docx::Document do
  let(:fixtures_path) { 'spec/fixtures' }
  let(:formatting_line_count) { 12 }

  describe 'reading' do
    let(:doc) { Docx::Document.open("#{fixtures_path}/basic.docx") }

    it 'should read the document' do
      expect(doc.paragraphs.size).to eq(2)
      expect(doc.paragraphs.first.text).to eq('hello')
      expect(doc.paragraphs.last.text).to eq('world')
      expect(doc.text).to eq("hello\nworld")
    end

    it 'should read bookmarks' do
      expect(doc.bookmarks.size).to eq(1)
      expect(doc.bookmarks['test_bookmark']).to_not eq(nil)
    end

    it 'should have paragraphs' do
      doc.each_paragraph do |p|
        expect(p).to be_an_instance_of(Docx::Elements::Containers::Paragraph)
      end
    end

    it 'should have properly formatted text runs' do
      doc.each_paragraph do |p|
        p.each_text_run do |tr|
          expect(tr).to be_an_instance_of(Docx::Elements::Containers::TextRun)
          expect(tr.formatting).to eq(Docx::Elements::Containers::TextRun::DEFAULT_FORMATTING)
        end
      end
    end
  end

  describe 'read tables' do
    let(:doc) { Docx::Document.open("#{fixtures_path}/tables.docx") }

    it "should have tables with rows and cells" do
      expect(doc.tables.count).to eq 2
      doc.tables.each do |table|
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
      doc.tables.each do |table|
        table.columns.each do |column|
          expect(column).to be_an_instance_of(Docx::Elements::Containers::TableColumn)
          column.cells.each do |cell|
            expect(cell).to be_an_instance_of(Docx::Elements::Containers::TableCell)
          end
        end
      end
    end

    it "should have proper count" do
      expect(doc.tables[0].row_count).to eq 171
      expect(doc.tables[1].row_count).to eq 2
      expect(doc.tables[0].column_count).to eq 2
      expect(doc.tables[1].column_count).to eq 2
    end

    it "should have tables with proper text" do
      expect(doc.tables[0].rows[0].cells[0].text).to eq "ENGLISH"
      expect(doc.tables[0].rows[0].cells[1].text).to eq "FRANÇAIS"
      expect(doc.tables[1].rows[0].cells[0].text).to eq "Second table"
      expect(doc.tables[1].rows[0].cells[1].text).to eq "Second tableau"
      expect(doc.tables[0].columns[0].cells[5].text).to eq "aphids"
      expect(doc.tables[0].columns[1].cells[5].text).to eq "puceron"
    end

    it "should read embedded links" do
      expect(doc.tables[0].columns[1].cells[1].text).to match(/^Directive/)
    end
  end

  describe 'editing'  do
    let (:doc) { Docx::Document.open("#{fixtures_path}/editing.docx") }

    it 'should copy paragraphs' do
      old_p = doc.paragraphs.first
      new_p = old_p.copy
      expect(new_p).to be_an_instance_of(Docx::Elements::Containers::Paragraph)
      expect(new_p).not_to eq(nil)
      expect(new_p).not_to eq(old_p)
    end

    it 'allows insertion of text' do
      expect(doc.paragraphs.size).to eq(3)
      first_p = doc.paragraphs.first
      new_p = first_p.copy
      new_p.insert_after first_p
      expect(doc.paragraphs.size).to eq(4)
    end

    it 'should change text' do
      expect(doc.paragraphs.first.text).to eq('test text')
      doc.paragraphs.first.text = 'the real test'
      expect(doc.paragraphs.first.text).to eq('the real test')
    end

    it 'should allow insertion of text before a bookmark' do
      expect(doc.paragraphs.first.text).to eq('test text')
      doc.bookmarks['beginning_bookmark'].insert_text_before('foo')
      expect(doc.paragraphs.first.text).to eq('footest text')
    end

    it 'should allow insertion of text after a bookmark' do
      expect(doc.paragraphs.first.text).to eq('test text')
      doc.bookmarks['end_bookmark'].insert_text_after('bar')
      expect(doc.paragraphs.first.text).to eq('test textbar')
    end

    it 'should allow multiple lines of text to be inserted at a bookmark' do
      expect(doc.paragraphs.last.text).to eq('')
      new_lines = ['replacement test', 'second paragraph test', 'and a third paragraph test']
      doc.bookmarks['isolated_bookmark'].insert_multiple_lines(new_lines)
      new_lines.each_index do |line|
        expect(doc.paragraphs[line + 2].text).to eq(new_lines[line])
      end
    end

    it 'should allow multi-line insertion with replacement' do
      expect(doc.paragraphs[1].text).to eq('placeholder text')
      new_lines = ['replacement test', 'second paragraph test', 'and a third paragraph test']
      doc.bookmarks['word_splitting_bookmark'].insert_multiple_lines(new_lines)
      new_lines.each_index do |line|
        expect(doc.paragraphs[line + 1].text).to eq(new_lines[line])
      end
    end

    it 'should allow content deletion' do
      expect(doc.paragraphs.first.text).to eq('test text')
      doc.paragraphs.first.blank!
      expect(doc.paragraphs.first.text).to eq('')
    end

    it 'should allow content deletion' do
      expect{doc.paragraphs.first.remove!}.to change{doc.paragraphs.size}.by(-1)

    end
  end

  describe 'read formatting' do
    let(:doc) do
      Docx::Document.open("#{fixtures_path}/formatting.docx")
    end
    let(:formatting) do
      doc.paragraphs.map { |p| p.text_runs.map(&:formatting) }
    end
    let(:default_formatting) do
      Docx::Elements::Containers::TextRun::DEFAULT_FORMATTING
    end
    let(:only_italic) { default_formatting.merge(italic: true) }
    let(:only_bold) { default_formatting.merge(bold: true) }
    let(:only_underline) { default_formatting.merge(underline: true) }
    let(:all_formatted) do
      default_formatting.merge(italic: true, bold: true, underline: true)
    end

    it 'should have the correct text' do
      expect(doc.paragraphs.size).to eq(formatting_line_count)
      expect(doc.paragraphs[0].text).to eq('Normal')
      expect(doc.paragraphs[1].text).to eq('Italic')
      expect(doc.paragraphs[2].text).to eq('Bold')
      expect(doc.paragraphs[3].text).to eq('Underline')
      expect(doc.paragraphs[4].text).to eq('Normal')
      expect(doc.paragraphs[5].text).to eq('This is a sentence with all formatting options in the middle of the sentence.')
      expect(doc.paragraphs[6].text).to eq('This is a centered paragraph.')
      expect(doc.paragraphs[7].text).to eq('This paragraph is aligned left.')
      expect(doc.paragraphs[8].text).to eq('This paragraph is aligned right.')
      expect(doc.paragraphs[9].text).to eq('This paragraph is 14 points.')
      expect(doc.paragraphs[10].text).to eq('This paragraph has a word at 16 points.')
    end

    it 'should contain a paragraph with multiple text runs' do

    end

    it 'should detect normal formatting' do
      [0, 4].each do |i|
        expect(formatting[i][0]).to eq(default_formatting)
        expect(doc.paragraphs[i].text_runs[0].italicized?).to eq(false)
        expect(doc.paragraphs[i].text_runs[0].bolded?).to eq(false)
        expect(doc.paragraphs[i].text_runs[0].underlined?).to eq(false)
      end
    end

    it 'should detect italic formatting' do
      expect(formatting[1][0]).to eq(only_italic)
      expect(doc.paragraphs[1].text_runs[0].italicized?).to eq(true)
      expect(doc.paragraphs[1].text_runs[0].bolded?).to eq(false)
      expect(doc.paragraphs[1].text_runs[0].underlined?).to eq(false)
    end

    it 'should detect bold formatting' do
      expect(formatting[2][0]).to eq(only_bold)
      expect(doc.paragraphs[2].text_runs[0].italicized?).to eq(false)
      expect(doc.paragraphs[2].text_runs[0].bolded?).to eq(true)
      expect(doc.paragraphs[2].text_runs[0].underlined?).to eq(false)
    end

    it 'should detect underline formatting' do
      expect(formatting[3][0]).to eq(only_underline)
      expect(doc.paragraphs[3].text_runs[0].italicized?).to eq(false)
      expect(doc.paragraphs[3].text_runs[0].bolded?).to eq(false)
      expect(doc.paragraphs[3].text_runs[0].underlined?).to eq(true)
    end

    it 'should detect mixed formatting' do
      expect(formatting[5][0]).to eq(default_formatting)
      expect(doc.paragraphs[5].text_runs[0].italicized?).to eq(false)
      expect(doc.paragraphs[5].text_runs[0].bolded?).to eq(false)
      expect(doc.paragraphs[5].text_runs[0].underlined?).to eq(false)
      
      expect(formatting[5][1]).to eq(all_formatted)
      expect(doc.paragraphs[5].text_runs[1].italicized?).to eq(true)
      expect(doc.paragraphs[5].text_runs[1].bolded?).to eq(true)
      expect(doc.paragraphs[5].text_runs[1].underlined?).to eq(true)
      
      expect(formatting[5][2]).to eq(default_formatting)
      expect(doc.paragraphs[5].text_runs[2].italicized?).to eq(false)
      expect(doc.paragraphs[5].text_runs[2].bolded?).to eq(false)
      expect(doc.paragraphs[5].text_runs[2].underlined?).to eq(false)
    end

    it 'should detect centered paragraphs' do
      expect(doc.paragraphs[5].aligned_center?).to eq(false)
      expect(doc.paragraphs[6].aligned_center?).to eq(true)
      expect(doc.paragraphs[7].aligned_center?).to eq(false)
    end

    it 'should detect left justified paragraphs' do
      expect(doc.paragraphs[6].aligned_left?).to eq(false)
      expect(doc.paragraphs[7].aligned_left?).to eq(true)
      expect(doc.paragraphs[8].aligned_left?).to eq(false)
    end

    it 'should detect right justified paragraphs' do
      expect(doc.paragraphs[7].aligned_right?).to eq(false)
      expect(doc.paragraphs[8].aligned_right?).to eq(true)
      expect(doc.paragraphs[9].aligned_right?).to eq(false)
    end

    # ECMA-376 Office Open XML spec (4th edition), 17.3.2.38, size is
    # defined in half-points, meaning 14pt text returns a value of 28.
    # http://www.ecma-international.org/publications/standards/Ecma-376.htm
    it 'should return proper font size for paragraphs' do
      expect(doc.font_size).to eq 11
      expect(doc.paragraphs[5].font_size).to eq 11
      paragraph = doc.paragraphs[9]
      expect(paragraph.font_size).to eq 14
      expect(paragraph.text_runs[0].font_size).to eq 14
    end

    it 'should return proper font size for runs' do
      expect(doc.font_size).to eq 11
      paragraph = doc.paragraphs[10]
      expect(paragraph.font_size).to eq 11
      text_runs = paragraph.text_runs
      expect(text_runs[0].font_size).to eq 11
      expect(text_runs[1].font_size).to eq 16
      expect(text_runs[2].font_size).to eq 11
      expect(text_runs[3].font_size).to eq 11
      expect(text_runs[4].font_size).to eq 11
    end
  end

  describe 'saving' do
    let (:doc) { Docx::Document.open("#{fixtures_path}/saving.docx") }
    let (:new_doc_path) { "#{fixtures_path}/new_save.docx" }

    it 'should save to a normal file path' do
      doc.save(new_doc_path)
      new_doc = Docx::Document.open(new_doc_path)
      expect(new_doc.paragraphs.size).to eq(doc.paragraphs.size)
    end

    context 'saving to a tempfile' do
      let(:temp_file) { Tempfile.new(['docx_gem', '.docx']) }
      let(:new_doc_path) { temp_file.path }
      it 'should save to a tempfile' do
        doc.save(new_doc_path)
        new_doc = Docx::Document.open(new_doc_path)
        expect(new_doc.paragraphs.size).to eq(doc.paragraphs.size)

        temp_file.close
        temp_file.unlink
        # ensure temp file has been removed
        expect(File.exists?(new_doc_path)).to eq(false)
      end
    end

    after do
      if File.exists?(new_doc_path)
        File.delete(new_doc_path)
      end
    end
  end

  describe 'outputting html' do
    let(:doc) { Docx::Document.open("#{fixtures_path}/formatting.docx") }

    let(:formatted_line) { doc.paragraphs[5] }
    let(:p_regex) { /(^\<p).+((?<=\>)\w+)(\<\/p>$)/ }
    let(:span_regex) { /(\<span).+((?<=\>)\w+)(<\/span>)/ }
    let(:em_regex) { /(\<em).+((?<=\>)\w+)(\<\/em\>)/ }
    let(:strong_regex) { /(\<strong).+((?<=\>)\w+)(\<\/strong\>)/ }

    it 'should wrap pragraphs in a p tag' do
      scan = doc.paragraphs[0].to_html.scan(p_regex).flatten
      expect(scan.first).to eq('<p')
      expect(scan.last).to eq('</p>')
      expect(scan[1]).to eq('Normal')
    end
   
    it 'should emphasize italicized text' do
      scan = doc.paragraphs[1].to_html.scan(em_regex).flatten
      expect(scan.first).to eq('<em')
      expect(scan.last).to eq('</em>')
      expect(scan[1]).to eq('Italic')
    end

    it 'should strong bolded text' do
      scan = doc.paragraphs[2].to_html.scan(strong_regex).flatten
      expect(scan.first).to eq '<strong'
      expect(scan.last).to eq '</strong>'
      expect(scan[1]).to eq 'Bold'
    end

    it 'should underline underlined text' do
      scan = doc.paragraphs[3].to_html.scan(/\<span\s+([^\>]+)/).flatten
      expect(scan.first).to eq 'style="text-decoration:underline;"'
    end

    it 'should justify paragraphs' do
      regex = /^<p[^\"]+.(?<=\")([^\"]+)/
      expect(doc.paragraphs[6].to_html.scan(regex).flatten.first.split(';').include?('text-align:center')).to eq(true)
      expect(doc.paragraphs[7].to_html.scan(regex).flatten.first.split(';').include?('text-align:left')).to eq(false)
      expect(doc.paragraphs[8].to_html.scan(regex).flatten.first.split(';').include?('text-align:right')).to eq(true)
    end

    it "should set font size on styled paragraphs" do
      regex = /(\<p{1})[^\>]+style\=\"([^\"]+).+(<\/p>)/      
      scan = doc.paragraphs[9].to_html.scan(regex).flatten
      expect(scan.first).to eq '<p'
      expect(scan.last).to eq '</p>'
      expect(scan[1].split(';').include?('font-size:14pt')).to eq(true)
    end

    it 'should set font size on styled text runs' do
      regex = /(\<span)[^\>]+style\=\"([^\"]+)[^\<]+(<\/span>)/
      scan = doc.paragraphs[10].to_html.scan(regex).flatten
      expect(scan.first).to eq '<span'
      expect(scan.last).to eq '</span>'
      expect(scan[1].split(';').include?('font-size:16pt')).to eq(true)
    end

    it 'should properly highlight different text in different places in a sentence' do
      paragraph = doc.paragraphs[11]
      scan = paragraph.to_html.scan(em_regex).flatten
      expect(scan.first).to eq '<em'
      expect(scan.last).to eq '</em>'
      expect(scan[1]).to eq 'sentence'
      scan = paragraph.to_html.scan(strong_regex).flatten
      expect(scan.first).to eq '<strong'
      expect(scan.last).to eq '</strong>'
      expect(scan[1]).to eq 'formatting'
      scan = paragraph.to_html.scan(span_regex).flatten
      expect(scan.first).to eq '<span'
      expect(scan.last).to eq '</span>'
      expect(scan[1]).to eq 'different'
      scan = paragraph.to_html.scan(/\<span\s+([^\>]+)/).flatten
      expect(scan.first).to eq 'style="text-decoration:underline;"'
    end

    it 'should output an entire document as html fragment' do
      expect(doc.to_html.scan(/(\<p)/).flatten.size).to eq(formatting_line_count)
    end

    it 'should output styled html' do
      expect(formatted_line.to_html.scan('<span style="text-decoration:underline;"><strong><em>all</em></strong></span>').size).to eq 1
    end

  end

  describe 'replacing contents' do
    let(:replacement_file_path) { "#{fixtures_path}/replacement.png" }
    let(:temp_file_path){ Tempfile.new(['docx_gem', '.docx']).path }
    let(:entry_path){ 'word/media/image1.png' }
    let(:doc){ Docx::Document.open("#{fixtures_path}/replacement.docx") }

    it 'should replace existing file within the document' do
      File.open replacement_file_path, "rb" do |io|
        doc.replace_entry entry_path, io.read
      end

      doc.save(temp_file_path)

      File.open replacement_file_path, "rb" do |io|
        expect(Zip::File.open(temp_file_path).read entry_path).to eq io.read
      end
    end

    after do
      if File.exists?(temp_file_path)
        File.delete(temp_file_path)
      end
    end
  end
end

