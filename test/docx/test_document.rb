require 'docx'
require 'test/unit'

class DocxTest < Test::Unit::TestCase
  def setup
    @doc = Docx::Document.open('test/fixtures/basic.docx')
  end
  
  def test_basic_functionality
    assert_equal 2, @doc.paragraphs.size
    assert_equal 'hello', @doc.paragraphs.first.to_s
    assert_equal 'world', @doc.paragraphs.last.to_s
    assert_equal "hello\nworld", @doc.to_s
  end
  
  def test_each_paragraph
    @doc.each_paragraph do |p|
      assert p.kind_of?(Docx::Containers::Paragraph)
      assert p.text_runs.all? { |r| r.formatting == :none }
    end
  end
end
