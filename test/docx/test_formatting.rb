require 'docx'
require 'test/unit'

class FormattingTest < Test::Unit::TestCase
  def setup
    @doc = Docx::Document.open('test/fixtures/formatting.docx')
    @formatting = @doc.paragraphs.map { |p| p.text_runs.map(&:formatting) }
  end
  
  def test_document_text
    assert_equal 6,           @doc.paragraphs.size
    assert_equal 'Normal',    @doc.paragraphs[0].text
    assert_equal 'Italic',    @doc.paragraphs[1].text
    assert_equal 'Bold',      @doc.paragraphs[2].text
    assert_equal 'Underline', @doc.paragraphs[3].text
    assert_equal 'Normal',    @doc.paragraphs[4].text
    assert_equal 'This is a sentence with all formatting options in the middle of the sentence.',
                              @doc.paragraphs[5].text
  end
  
  def test_multiple_text_runs_in_paragraph
    p = @doc.paragraphs[5]
    assert_equal 3, p.text_runs.size
    assert_equal 'This is a sentence with ', p.text_runs[0].text
    assert_equal 'all', p.text_runs[1].text
    assert_equal ' formatting options in the middle of the sentence.', p.text_runs[2].text
  end
  
  def test_normal_formatting
    [0, 4].each do |i|
      assert_equal default_formatting, @formatting[i][0]
      refute @doc.paragraphs[i].text_runs[0].italicized?
      refute @doc.paragraphs[i].text_runs[0].bolded?
      refute @doc.paragraphs[i].text_runs[0].underlined?
    end
  end
  
  def test_italic_formatting
    assert_equal only_italic, @formatting[1][0]
    assert @doc.paragraphs[1].text_runs[0].italicized?
    refute @doc.paragraphs[1].text_runs[0].bolded?
    refute @doc.paragraphs[1].text_runs[0].underlined?
  end
  
  def test_bold_formatting
    assert_equal only_bold, @formatting[2][0]
    refute @doc.paragraphs[2].text_runs[0].italicized?
    assert @doc.paragraphs[2].text_runs[0].bolded?
    refute @doc.paragraphs[2].text_runs[0].underlined?
  end
  
  def test_underline_formatting
    assert_equal only_underline, @formatting[3][0]
    refute @doc.paragraphs[3].text_runs[0].italicized?
    refute @doc.paragraphs[3].text_runs[0].bolded?
    assert @doc.paragraphs[3].text_runs[0].underlined?
  end
  
  def test_mixed_formatting
    assert_equal default_formatting, @formatting[5][0]
    refute @doc.paragraphs[5].text_runs[0].italicized?
    refute @doc.paragraphs[5].text_runs[0].bolded?
    refute @doc.paragraphs[5].text_runs[0].underlined?
    
    assert_equal all_formatting_options, @formatting[5][1]
    assert @doc.paragraphs[5].text_runs[1].italicized?
    assert @doc.paragraphs[5].text_runs[1].bolded?
    assert @doc.paragraphs[5].text_runs[1].underlined?
    
    assert_equal default_formatting, @formatting[5][2]
    refute @doc.paragraphs[5].text_runs[2].italicized?
    refute @doc.paragraphs[5].text_runs[2].bolded?
    refute @doc.paragraphs[5].text_runs[2].underlined?
  end
  
  private
  
  def default_formatting
    Docx::Containers::TextRun::DEFAULT_FORMATTING
  end
  
  def only_italic
    default_formatting.merge italic: true
  end
  
  def only_bold
    default_formatting.merge bold: true
  end
  
  def only_underline
    default_formatting.merge underline: true
  end
  
  def all_formatting_options
    default_formatting.merge italic: true, bold: true, underline: true
  end
end
