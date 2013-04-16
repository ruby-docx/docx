# docx

a ruby library/gem for interacting with `.docx` files. currently capabilities include reading paragraphs/bookmarks, inserting text at bookmarks, and saving the document.

## usage

### install

requires ruby (only tested with 1.9.3 so far)

    gem install docx

### reading

``` ruby
require 'docx'

d = Docx::Document.open('example.docx')
# Array of paragraphs
d.paragraphs.each do |p|
  puts d
end

# Hash of Bookmarks. Bookmark names as keys correspond to bookmark objects.
d.bookmarks.each_pair do |bookmark_name, bookmark_object|
  puts bookmark_name
end
```

### writing

``` ruby
require 'docx'

d = Docx::Document.open('example.docx')
# Insert a single line after a bookmark
d.bookmarks['example_bookmark'].insert_after("Hello world.")
# Each value in array is put on a separate line
d.bookmarks['example_bookmark'].insert_multiple_lines_after(['Hello', 'World', 'foo'])
```

### advanced

``` ruby
require 'docx'

d = Docx::Document.open('example.docx')

# The Nokogiri node on which an element is based can be accessed using #node
d.paragraphs.each do |p|
  puts p.node.inspect
end

# The #xpath and #at_xpath are delegated to the node from the element, saving a step
p_element = d.paragraphs.first
p_children = p_element.xpath("//child::*") # selects all children
p_child = p_element.at_xpath("//child::*") # selects first child
```

## Development

### todo

* Add better formatting identification for specific nodes and other formatting indicators (text size, paragraph spacing)
* Calculate element formatting based on values present in element properties as well as properties inherited from parents
* Default formatting of inserted elements to inherited values
* Implement formattable elements.
* Implement styles.
* Easier multi-line text insertion at a single bookmark (inserting paragraph nodes after the one containing the bookmark)