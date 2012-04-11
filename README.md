# docx

a ruby library/gem for interacting with `.docx` files

## usage

### basic

``` ruby
require 'docx'

d = Docx::Document.open('example.docx')
d.each_paragraph do |p|
  puts d
end
```

### advanced

``` ruby
require 'docx'

d = Docx::Document.open('example.docx')
d.each_paragraph do |p|
  p.each_text_run do |run|
    run.italicized?
    run.bolded?
    run.underlined?
    run.formatting
    run.text
  end
end
