require 'docx'
require 'test/unit'

class SaveTest < Test::Unit::TestCase
  def setup
    @doc = Docx::Document.open('test/fixtures/saving.docx')
  end
end
