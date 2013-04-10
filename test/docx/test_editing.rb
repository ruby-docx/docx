require 'docx'
require 'test/unit'

class EditTest < Test::Unit::TestCase
  def setup
    @doc = Docx::Document.open('test/fixtures/editing.docx')
  end

  def test_copy_paragraph
    old_p = @doc.paragraphs.first
    new_p = old_p.copy
    assert_kind_of Docx::Elements::Containers::Paragraph, new_p
    assert_not_nil new_p
    assert_not_same old_p, new_p
  end

  def test_insertion
    assert_equal 3, @doc.paragraphs.size
    first_p = @doc.paragraphs.first
    new_p = first_p.copy
    new_p.insert_after first_p
    assert_equal 4, @doc.paragraphs.size
  end

  def test_paragraph_text
    assert_equal 'test text', @doc.paragraphs.first.text
    @doc.paragraphs.first.text = 'the real test'
    assert_equal 'the real test', @doc.paragraphs.first.text
  end

  def test_inserting_text_before_bookmark
    assert_equal 'test text', @doc.paragraphs.first.text
    @doc.bookmarks['beginning_bookmark'].insert_text_before('foo')
    assert_equal 'footest text', @doc.paragraphs.first.text
  end

  def test_inserting_text_after_bookmark
    assert_equal 'test text', @doc.paragraphs.first.text
    @doc.bookmarks['end_bookmark'].insert_text_after('bar')
    assert_equal 'test textbar', @doc.paragraphs.first.text
  end

  def test_inserting_multiple_lines_at_bookmark
    assert_equal '', @doc.paragraphs.last.text
    new_lines = ['replacement test', 'second paragraph test', 'and a third paragraph test']
    @doc.bookmarks['isolated_bookmark'].insert_multiple_lines(new_lines)
    new_lines.each_index do |line|
      assert_equal new_lines[line], @doc.paragraphs[line + 2].text
    end
  end

  def test_inserting_multiple_lines_at_bookmark_with_replacement
    assert_equal 'placeholder text', @doc.paragraphs[1].text
    new_lines = ['replacement test', 'second paragraph test', 'and a third paragraph test']
    @doc.bookmarks['word_splitting_bookmark'].insert_multiple_lines(new_lines)
    new_lines.each_index do |line|
      assert_equal new_lines[line], @doc.paragraphs[line + 1].text
    end
  end

  # Insert text intelligently around bookmark
  def test_inserting_text_around_bookmark

  end

  def test_blank
    assert_equal 'test text', @doc.paragraphs.first.text
    @doc.paragraphs.first.blank!
    assert_equal '', @doc.paragraphs.first.text
  end
end
