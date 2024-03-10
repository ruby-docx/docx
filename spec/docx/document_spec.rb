# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'tempfile'

describe Docx::Document do
  before(:all) do
    @fixtures_path = 'spec/fixtures'
    @formatting_line_count = 15 # number of lines the formatting.docx file has
  end

  describe '#open' do
    context 'When reading a file made by Office365' do
      it 'supports it' do
        expect do
          Docx::Document.open(@fixtures_path + '/office365.docx')
        end.to_not raise_error
      end
    end

    context 'When reading a un-supported file' do
      it 'should throw file not supported error' do
        expect do
          Docx::Document.open(@fixtures_path + '/invalid_format.pdf')
        end.to raise_error(Errno::EIO, 'Input/output error - Invalid file format')
      end

      it 'should throw file not found error' do
        invalid_path = @fixtures_path + '/invalid_file_path.docx'
        expect do
          Docx::Document.open(invalid_path)
        end.to raise_error(Zip::Error, "File #{invalid_path} not found")
      end
    end
  end

  describe "#inspect" do
    it "isn't too long" do
      doc = Docx::Document.open(@fixtures_path + '/office365.docx')

      expect(doc.inspect.length).to be < 1000

      doc.instance_variables.each do |var|
        expect(doc.inspect).to match(/#{var}/)
      end
    end
  end

  describe 'reading' do
    context 'using normal file' do
      before do
        @doc = Docx::Document.open(@fixtures_path + '/basic.docx')
      end

      it_behaves_like 'reading'
    end

    context 'using stream' do
      before do
        stream = File.binread(@fixtures_path + '/basic.docx')
        @doc = Docx::Document.open(stream)
      end

      it_behaves_like 'reading'
    end
  end

  describe 'read tables' do
    before do
      @doc = Docx::Document.open(@fixtures_path + '/tables.docx')
    end

    it 'should have tables with rows and cells' do
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

    it 'should have tables with columns and cells' do
      @doc.tables.each do |table|
        table.columns.each do |column|
          expect(column).to be_an_instance_of(Docx::Elements::Containers::TableColumn)
          column.cells.each do |cell|
            expect(cell).to be_an_instance_of(Docx::Elements::Containers::TableCell)
          end
        end
      end
    end

    it 'should have proper count' do
      expect(@doc.tables[0].row_count).to eq 171
      expect(@doc.tables[1].row_count).to eq 2
      expect(@doc.tables[0].column_count).to eq 2
      expect(@doc.tables[1].column_count).to eq 2
    end

    it 'should have tables with proper text' do
      expect(@doc.tables[0].rows[0].cells[0].text).to eq 'ENGLISH'
      expect(@doc.tables[0].rows[0].cells[1].text).to eq 'FRANÇAIS'
      expect(@doc.tables[1].rows[0].cells[0].text).to eq 'Second table'
      expect(@doc.tables[1].rows[0].cells[1].text).to eq 'Second tableau'
      expect(@doc.tables[0].columns[0].cells[5].text).to eq 'aphids'
      expect(@doc.tables[0].columns[1].cells[5].text).to eq 'puceron'
    end

    it 'should read embedded links' do
      expect(@doc.tables[0].columns[1].cells[1].text).to match(/^Directive/)
    end

    describe '#paragraphs' do
      it 'should not grabs paragraphs in the tables' do
        expect(@doc.paragraphs.map(&:text)).to_not include("Second table")
      end
    end
  end

  describe 'editing' do
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

    it 'should allow content deletion' do
      expect { @doc.paragraphs.first.remove! }.to change { @doc.paragraphs.size }.by(-1)
    end
  end

  describe 'format-preserving substitution' do
    before do
      @doc = Docx::Document.open(@fixtures_path + '/substitution.docx')
    end

    it 'should replace placeholder in any line of a paragraph' do
      expect(@doc.paragraphs[0].text).to eq('Page title')
      expect(@doc.paragraphs[1].text).to eq('Multi-line paragraph line 1_placeholder2_ line 2_placeholder3_ line3 ')

      @doc.paragraphs[1].each_text_run do |text_run|
        text_run.substitute('_placeholder2_', 'same paragraph')
        text_run.substitute('_placeholder3_', 'yet the same paragraph')
      end

      expect(@doc.paragraphs[1].text).to eq('Multi-line paragraph line 1same paragraph line 2yet the same paragraph line3 ')
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
      expect(@doc.paragraphs.size).to eq(@formatting_line_count)
      expect(@doc.paragraphs[0].text).to eq('Normal')
      expect(@doc.paragraphs[1].text).to eq('Italic')
      expect(@doc.paragraphs[2].text).to eq('Bold')
      expect(@doc.paragraphs[3].text).to eq('Underline')
      expect(@doc.paragraphs[4].text).to eq('Normal')
      expect(@doc.paragraphs[5].text).to eq('This is a sentence with all formatting options in the middle of the sentence.')
      expect(@doc.paragraphs[6].text).to eq('This is a centered paragraph.')
      expect(@doc.paragraphs[7].text).to eq('This paragraph is aligned left.')
      expect(@doc.paragraphs[8].text).to eq('This paragraph is aligned right.')
      expect(@doc.paragraphs[9].text).to eq('This paragraph is 14 points.')
      expect(@doc.paragraphs[10].text).to eq('This paragraph has a word at 16 points.')
      expect(@doc.paragraphs[11].text).to eq('This sentence has different formatting in different places.')
      expect(@doc.paragraphs[12].text).to eq('This sentence has a hyperlink.')
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

    it 'should detect centered paragraphs' do
      expect(@doc.paragraphs[5].aligned_center?).to eq(false)
      expect(@doc.paragraphs[6].aligned_center?).to eq(true)
      expect(@doc.paragraphs[7].aligned_center?).to eq(false)
    end

    it 'should detect left justified paragraphs' do
      expect(@doc.paragraphs[6].aligned_left?).to eq(false)
      expect(@doc.paragraphs[7].aligned_left?).to eq(true)
      expect(@doc.paragraphs[8].aligned_left?).to eq(false)
    end

    it 'should detect right justified paragraphs' do
      expect(@doc.paragraphs[7].aligned_right?).to eq(false)
      expect(@doc.paragraphs[8].aligned_right?).to eq(true)
      expect(@doc.paragraphs[9].aligned_right?).to eq(false)
    end

    # ECMA-376 Office Open XML spec (4th edition), 17.3.2.38, size is
    # defined in half-points, meaning 14pt text returns a value of 28.
    # http://www.ecma-international.org/publications/standards/Ecma-376.htm
    it 'should return proper font size for paragraphs' do
      expect(@doc.font_size).to eq 11
      expect(@doc.paragraphs[5].font_size).to eq 11
      paragraph = @doc.paragraphs[9]
      expect(paragraph.font_size).to eq 14
      expect(paragraph.text_runs[0].font_size).to eq 14
    end

    it 'should return proper font size for runs' do
      expect(@doc.font_size).to eq 11
      paragraph = @doc.paragraphs[10]
      expect(paragraph.font_size).to eq 11
      text_runs = paragraph.text_runs
      expect(text_runs[0].font_size).to eq 11
      expect(text_runs[1].font_size).to eq 16
      expect(text_runs[2].font_size).to eq 11
      expect(text_runs[3].font_size).to eq 11
      expect(text_runs[4].font_size).to eq 11
    end

    it 'should return changed value for runs' do
      paragraph = @doc.paragraphs[10]
      text_runs = paragraph.text_runs

      tr = text_runs[0]
      expect(tr.text).to eq 'This paragraph has a '

      tr.text = 'This paragraph hasn\'t a'
      expect(tr.text).to eq 'This paragraph hasn\'t a'
    end
  end

  describe 'saving' do
    context 'from a normal file' do
      before do
        @doc = Docx::Document.open(@fixtures_path + '/saving.docx')
      end

      it_behaves_like 'saving to file'
    end

    context 'from a stream' do
      before do
        stream = File.binread(@fixtures_path + '/saving.docx')
        @doc = Docx::Document.open(stream)
      end

      it_behaves_like 'saving to file'
    end

    context 'wps modified docx file' do
      before { @doc = Docx::Document.open(@fixtures_path + '/saving_wps.docx') }

      it 'should save to a normal file path' do
        @new_doc_path = @fixtures_path + '/new_save.docx'
        @doc.save(@new_doc_path)
        @new_doc = Docx::Document.open(@new_doc_path)
        expect(@new_doc.paragraphs.size).to eq(@doc.paragraphs.size)
      end

      after { File.delete(@new_doc_path) if File.exist?(@new_doc_path) }
    end
  end

  describe 'streaming' do
    it 'should return a StringIO to send over HTTP' do
      doc = Docx::Document.open(@fixtures_path + '/basic.docx')
      expect(doc.stream).to be_a(StringIO)
    end

    context 'should return a valid docx stream' do
      before do
        doc = Docx::Document.open(@fixtures_path + '/basic.docx')
        result = doc.stream

        @doc = Docx::Document.open(result)
      end

      it_behaves_like 'reading'
    end
  end

  describe 'outputting html' do
    before do
      @doc = Docx::Document.open(@fixtures_path + '/formatting.docx')
      @formatted_line = @doc.paragraphs[5]
      @p_regex = /(^\<p).+((?<=\>)\w+)(\<\/p>$)/
      @span_regex = /(\<span).+((?<=\>)\w+)(<\/span>)/
      @em_regex = /(\<em).+((?<=\>)\w+)(\<\/em\>)/
      @strong_regex = /(\<strong).+((?<=\>)\w+)(\<\/strong\>)/
      @strike_regex = /(\<s).+((?<=\>)\w+)(\<\/s\>)/
      @anchor_tag_regex = /\<a href="(.+)" target="_blank"\>(.+)\<\/a>/
    end

    it 'should wrap pragraphs in a p tag' do
      scan = @doc.paragraphs[0].to_html.scan(@p_regex).flatten
      expect(scan.first).to eq('<p')
      expect(scan.last).to eq('</p>')
      expect(scan[1]).to eq('Normal')
    end

    it 'should emphasize italicized text' do
      scan = @doc.paragraphs[1].to_html.scan(@em_regex).flatten
      expect(scan.first).to eq('<em')
      expect(scan.last).to eq('</em>')
      expect(scan[1]).to eq('Italic')
    end

    it 'should strong bolded text' do
      scan = @doc.paragraphs[2].to_html.scan(@strong_regex).flatten
      expect(scan.first).to eq '<strong'
      expect(scan.last).to eq '</strong>'
      expect(scan[1]).to eq 'Bold'
    end

    it 'should underline underlined text' do
      scan = @doc.paragraphs[3].to_html.scan(/\<span\s+([^\>]+)/).flatten
      expect(scan.first).to eq 'style="text-decoration:underline;"'
    end

    it 'should strike striked text' do
      scan = @doc.paragraphs[13].to_html.scan(@strike_regex).flatten
      expect(scan.first).to eq '<s'
      expect(scan.last).to eq '</s>'
      expect(scan[1]).to eq 'Strike'
    end

    it 'should color the text' do
      scan = @doc.paragraphs[14].to_html.scan(/\<p\s+([^\>]+)/).flatten
      expect(scan.first).to eq 'style="font-size:11pt;color:#FF0000;"'
    end

    it 'should justify paragraphs' do
      regex = /^<p[^\"]+.(?<=\")([^\"]+)/
      expect(@doc.paragraphs[6].to_html.scan(regex).flatten.first.split(';').include?('text-align:center')).to eq(true)
      expect(@doc.paragraphs[7].to_html.scan(regex).flatten.first.split(';').include?('text-align:left')).to eq(false)
      expect(@doc.paragraphs[8].to_html.scan(regex).flatten.first.split(';').include?('text-align:right')).to eq(true)
    end

    it 'should set font size on styled paragraphs' do
      regex = /(\<p{1})[^\>]+style\=\"([^\"]+).+(<\/p>)/
      scan = @doc.paragraphs[9].to_html.scan(regex).flatten
      expect(scan.first).to eq '<p'
      expect(scan.last).to eq '</p>'
      expect(scan[1].split(';').include?('font-size:14pt')).to eq(true)
    end

    it 'should set font size on styled text runs' do
      regex = /(\<span)[^\>]+style\=\"([^\"]+)[^\<]+(<\/span>)/
      scan = @doc.paragraphs[10].to_html.scan(regex).flatten
      expect(scan.first).to eq '<span'
      expect(scan.last).to eq '</span>'
      expect(scan[1].split(';').include?('font-size:16pt')).to eq(true)
    end

    it 'should properly highlight different text in different places in a sentence' do
      paragraph = @doc.paragraphs[11]
      scan = paragraph.to_html.scan(@em_regex).flatten
      expect(scan.first).to eq '<em'
      expect(scan.last).to eq '</em>'
      expect(scan[1]).to eq 'sentence'
      scan = paragraph.to_html.scan(@strong_regex).flatten
      expect(scan.first).to eq '<strong'
      expect(scan.last).to eq '</strong>'
      expect(scan[1]).to eq 'formatting'
      scan = paragraph.to_html.scan(@span_regex).flatten
      expect(scan.first).to eq '<span'
      expect(scan.last).to eq '</span>'
      expect(scan[1]).to eq 'different'
      scan = paragraph.to_html.scan(/\<span\s+([^\>]+)/).flatten
      expect(scan.first).to eq 'style="text-decoration:underline;"'
    end

    it 'should output an entire document as html fragment' do
      expect(@doc.to_html.scan(/(\<p)/).flatten.size).to eq(@formatting_line_count)
    end

    it 'should output styled html' do
      expect(@formatted_line.to_html.scan('<span style="text-decoration:underline;"><strong><em>all</em></strong></span>').size).to eq 1
    end

    it 'should join paragraphs with newlines' do
      expect(@doc.to_html.scan(%(<p style="font-size:11pt;">Normal</p>\n<p style="font-size:11pt;"><em>Italic</em></p>\n<p style="font-size:11pt;"><strong>Bold</strong></p>)).size).to eq 1
    end

    it 'should convert hyperlinks to anchor tags' do
      scan = @doc.to_html.scan(@anchor_tag_regex).flatten
      expect(scan[0]).to eq "http://www.google.com/"
      expect(scan[1]).to eq "hyperlink"
    end
  end

  describe 'replacing contents' do
    let(:replacement_file_path) { @fixtures_path + '/replacement.png' }
    let(:temp_file_path) { Tempfile.new(['docx_gem', '.docx']).path }
    let(:entry_path) { 'word/media/image1.png' }
    let(:doc) { Docx::Document.open(@fixtures_path + '/replacement.docx') }

    it 'should replace existing file within the document' do
      File.open replacement_file_path, 'rb' do |io|
        doc.replace_entry entry_path, io.read
      end

      doc.save(temp_file_path)

      File.open replacement_file_path, 'rb' do |io|
        expect(Zip::File.open(temp_file_path).read(entry_path)).to eq io.read
      end
    end

    after do
      File.delete(temp_file_path) if File.exist?(temp_file_path)
    end
  end

  describe '#to_s' do
    let(:doc) { Docx::Document.open(@fixtures_path + '/weird_docx.docx') }

    it 'does not raise error' do
      expect { doc.to_s }.to_not raise_error
    end
    it 'returns a String' do
      expect(doc.to_s).to be_a(String)
    end
  end

  describe 'reading and manipulating paragraph style' do
    before do
      @doc = Docx::Document.open(@fixtures_path + '/styles.docx')
    end

    it 'read default style when not' do
      nb = @doc.paragraphs.size

      expect(@doc.paragraphs.map(&:style)).to eq([
        "Title",
        "Subtitle",
        "Author",
        "Date",
        "Compact",
        "Heading 1",
        "Heading 2",
        "Heading 3",
        "Heading 4",
        "Heading 5",
        "Heading 6",
        "Heading 7",
        "Heading 8",
        "Heading 9",
        "First Paragraph",
        "Body Text",
        "Block Text",
        "Table Caption",
        "Image Caption",
        "Definition Term",
        "Definition",
        "Definition Term",
        "Definition",
      ])

      expect(@doc.paragraphs.map(&:style_id)).to eq([
        "Title",
        "Subtitle",
        "Author",
        "Date",
        "Compact",
        "Heading1",
        "Heading2",
        "Heading3",
        "Heading4",
        "Heading5",
        "Heading6",
        "Heading7",
        "Heading8",
        "Heading9",
        "FirstParagraph",
        "BodyText",
        "BlockText",
        "TableCaption",
        "ImageCaption",
        "DefinitionTerm",
        "Definition",
        "DefinitionTerm",
        "Definition",
      ])
    end

    it 'set paragraph style' do
      nb = @doc.paragraphs.size
      expect(nb).to eq 23

      @doc.paragraphs.each do |p|
        p.style = 'Heading 1'
        expect(p.style).to eq 'Heading 1'
      end

      @doc.paragraphs.each do |p|
        p.style_id = 'Heading2'
        expect(p.style).to eq 'Heading 2'
      end
    end

    it 'raises if invalid paragraph style' do
      expect { @doc.paragraphs.first.style = 'invalid' }.to raise_error(Docx::Errors::StyleNotFound)
    end
  end

  describe 'reading and manipulating document styles' do
    before do
      @doc = Docx::Document.open(@fixtures_path + '/styles.docx')
    end

    it '#default_paragraphy_style' do
      expect(@doc.default_paragraph_style).to eq 'Normal'
    end

    it 'manipulates existing document styles' do
      styles_config = @doc.styles_configuration

      expect(styles_config.size).to eq 37

      heading_style = styles_config.style_of('Normal')
      expect(heading_style).to be_a(Docx::Elements::Style)

      expect(heading_style.id).to eq "Normal"
      expect(heading_style.font_color).to eq(nil)

      heading_style.font_color = "000000"
      expect(heading_style.font_color).to eq("000000")

      expect(heading_style.node.at_xpath("w:rPr/w:color/@w:val").value).to eq("000000")
    end

    it 'creates document styles' do
      styles_config = @doc.styles_configuration

      expect(styles_config.size).to eq 37
      expect { styles_config.style_of('Red') } .to raise_error(Docx::Errors::StyleNotFound)

      red_style = styles_config.add_style("Red")
      expect(styles_config.size).to eq 38

      expect(red_style).to be_a(Docx::Elements::Style)
      expect(red_style.id).to eq "Red"
      expect(red_style.name).to eq "Red"

      expect { red_style.font_color = "#FFFFFF" }.to raise_error(Docx::Errors::StyleInvalidPropertyValue)
      expect { red_style.font_color = "blue" }.to raise_error(Docx::Errors::StyleInvalidPropertyValue)
      expect { red_style.font_color = "FF0000" }.not_to raise_error

      styles_config.remove_style("Red")
      expect(styles_config.size).to eq 37
      expect { styles_config.style_of('Red') }.to raise_error(Docx::Errors::StyleNotFound)
    end

    it 'persists document styles' do
      styles_config = @doc.styles_configuration
      styles_config.add_style("Red", name: "Red", font_color: "FF0000", font_size: 20)
      @doc.paragraphs[5].style = "Red"

      first_modified_styles_path = @fixtures_path + '/styles_modified.docx'
      second_modified_styles_path = @fixtures_path + '/styles_modified2.docx'
      @doc.save(first_modified_styles_path)

      modified_styles_doc = Docx::Document.open(first_modified_styles_path)
      modified_styles_config = modified_styles_doc.styles_configuration

      expect(modified_styles_config.style_of('Red')).to be_a(Docx::Elements::Style)
      modified_styles_config.remove_style("Red")
      modified_styles_doc.save(second_modified_styles_path)

      modified_styles_doc = Docx::Document.open(second_modified_styles_path)
      modified_styles_config = modified_styles_doc.styles_configuration
      expect { modified_styles_config.style_of('Red') }.to raise_error(Docx::Errors::StyleNotFound)

      File.delete(first_modified_styles_path)
      File.delete(second_modified_styles_path)
    end

    after { File.delete(@new_doc_path) if @new_doc_path && File.exist?(@new_doc_path) }
  end

  describe '#to_html' do
    before do
      @doc = Docx::Document.open(@fixtures_path + '/internal-links.docx')
    end

    it 'should not raise error' do
      expect { @doc.to_html }.to_not raise_error
    end
  end
end
