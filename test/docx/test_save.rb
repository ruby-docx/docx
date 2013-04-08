require 'docx'
require 'test/unit'

class DocxTest < Test::Unit::TestCase
  def setup
    @doc = Docx::Document.open('test/fixtures/before_save.docx')
  end
  
  def test_basic_functionality
    
  end
  
  def test_each_paragraph
    
  end
  
  def test_each_text_run
    
  end
  
  private
  
  def default_formatting
    Docx::Containers::TextRun::DEFAULT_FORMATTING
  end
end