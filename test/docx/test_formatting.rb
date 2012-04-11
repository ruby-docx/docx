require 'docx'
require 'test/unit'

class FormattingTest < Test::Unit::TestCase
  def setup
    @doc = Docx::Document.open('test/fixtures/formatting.docx')
    @formatting = @doc.paragraphs.map { |p| p.text_runs.first.formatting }
  end
  
  def test_document_text
    assert_equal 5,           @doc.paragraphs.size
    assert_equal 'Normal',    @doc.paragraphs[0].text
    assert_equal 'Italic',    @doc.paragraphs[1].text
    assert_equal 'Bold',      @doc.paragraphs[2].text
    assert_equal 'Underline', @doc.paragraphs[3].text
    assert_equal 'normal',    @doc.paragraphs[4].text
  end
  
  def test_normal_formatting
    assert_equal default_formatting, @formatting[0]
    assert_equal default_formatting, @formatting[4]
  end
  
  def test_italic_formatting
    assert_equal only_italic, @formatting[1]
  end
  
  def test_bold_formatting
    assert_equal only_bold, @formatting[2]
  end
  
  def test_underline_formatting
    assert_equal only_underline, @formatting[3]
  end
  
  private
  
  def default_formatting
    Docx::Containers::TextRun::DEFAULT_FORMATTING
  end
  
  def only_italic
    {
      italic:    true,
      bold:      false,
      underline: false
    }
  end
  
  def only_bold
    {
      italic:    false,
      bold:      true,
      underline: false
    }
  end
  
  def only_underline
    {
      italic:    false,
      bold:      false,
      underline: true
    }
  end
end
