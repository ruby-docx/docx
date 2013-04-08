require_relative 'document'

document = Docx::Document.new('../test/fixtures/basic.docx')
puts document.bookmarks