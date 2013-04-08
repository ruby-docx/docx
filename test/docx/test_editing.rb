require 'docx'
require 'test/unit'

class SaveTest < Test::Unit::TestCase
  def setup
    @doc = Docx::Document.open('test/fixtures/saving.docx')
  end

  def test_copy
    old_p = @doc.paragraphs.first
    new_p = old_p.copy
    assert_kind_of Docx::Elements::Containers::Paragraph, new_p
    assert_not_nil new_p
    assert_not_same old_p, new_p
  end

  def test_insertion
    assert_equal 2, @doc.paragraphs.size
    first_p = @doc.paragraphs.first
    new_p = first_p.copy
    new_p.insert_after first_p
    assert_equal 3, @doc.paragraphs.size
  end

  def test_blank
    assert_equal 'test text', @doc.paragraphs.first.text
    @doc.paragraphs.first.blank!
    assert_equal '', @doc.paragraphs.first.text
  end
end
