require 'docx'
require 'test/unit'

class DocxTest < Test::Unit::TestCase
  def setup
    @doc = Docx::Document.open('test/fixtures/basic.docx')
  end
  
  def test_basic_functionality
    assert_equal "hello\nworld", @doc.to_s
  end
  
  def test_each_paragraph
    @doc.each_paragraph { |p| assert p.kind_of?(Docx::Containers::Paragraph) }
  end
end
