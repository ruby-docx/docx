require 'docx'
require 'test/unit'
require 'tempfile'

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

  def test_saving_to_tempfile
    temp_file = Tempfile.new(['docx_gem', '.docx'])
    @new_doc_path = temp_file.path
    @doc.save(@new_doc_path)
    @new_doc = Docx::Document.open(@new_doc_path)
    assert_equal @doc.paragraphs.size, @new_doc.paragraphs.size

    temp_file.close
    temp_file.unlink
    # ensure temp file has been removed
    assert_equal(false, File.exists?(@new_doc_path))
  end

  def teardown
    if File.exists?(@new_doc_path)
      File.delete(@new_doc_path)
    end
  end
end
