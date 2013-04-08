require 'docx'
require 'test/unit'

class SaveTest < Test::Unit::TestCase
  def setup
    @doc = Docx::Document.open('test/fixtures/saving.docx')
  end

  def test_saving
    @new_doc_path = 'test/fixtures/new_save.docx'
    @doc.save('test/fixtures/new_save.docx')
    @new_doc = Docx::Document.open('test/fixtures/saving.docx')
    assert_equal @doc.paragraphs.size, @new_doc.paragraphs.size
  end

  def teardown
    if @new_doc and @new_doc_path
      File.delete(@new_doc_path)
    end
  end
end
