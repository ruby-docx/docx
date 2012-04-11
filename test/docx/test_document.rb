require 'docx'
require 'test/unit'

class DocxTest < Test::Unit::TestCase
  def setup
    @doc = Docx::Document.open('test/fixtures/basic.docx')
  end
  
  def test_basic_functionality
    assert_equal 2, @doc.paragraphs.size
    assert_equal 'hello', @doc.paragraphs.first.text
    assert_equal 'world', @doc.paragraphs.last.text
    assert_equal "hello\nworld", @doc.text
  end
  
  def test_each_paragraph
    @doc.each_paragraph do |p|
      assert p.kind_of?(Docx::Containers::Paragraph)
      assert p.text_runs.all? { |r| r.formatting == Docx::Containers::TextRun::DEFAULT_FORMATTING }
    end
  end
end
