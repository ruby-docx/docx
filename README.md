# docx

a ruby library/gem for interacting with `.docx` files

## usage

### install

requires ruby (only tested with 1.9.3 so far)

    gem install docx

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
```

## Development

### todo

* Add better formatting identification for specific nodes and other formatting indicators (text size, paragraph spacing)
* Calculate element formatting based on values present in element properties as well as properties inherited from parents
* Default formatting of inserted elements to inherited values
* Implement formattable elements.
* Implement styles.
* Easier multi-line text insertion at a single bookmark (inserting paragraph nodes after the one containing the bookmark)