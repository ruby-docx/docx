# Docx

## TODO

### Basic Functionality

    require 'docx'
    
    d = Docx::Document.open('test.docx')
    d.each_paragraph do |p|
      puts d
    end
