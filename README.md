# docx

[![Gem Version](https://badge.fury.io/rb/docx.svg)](https://badge.fury.io/rb/docx)
[![Ruby](https://github.com/ruby-docx/docx/workflows/Ruby/badge.svg)](https://github.com/ruby-docx/docx/actions?query=workflow%3ARuby)
[![Coverage Status](https://coveralls.io/repos/github/ruby-docx/docx/badge.svg?branch=master)](https://coveralls.io/github/ruby-docx/docx?branch=master)
[![Gitter](https://badges.gitter.im/ruby-docx/community.svg)](https://gitter.im/ruby-docx/community?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge)

A ruby library/gem for interacting with `.docx` files. currently capabilities include reading paragraphs/bookmarks, inserting text at bookmarks, reading tables/rows/columns/cells and saving the document.

## Usage

### Prerequisites

- Ruby 2.6 or later

### Install

Add the following line to your application's Gemfile:

```ruby
gem 'docx'
```

And then execute:

```shell
bundle install
```

Or install it yourself as:

```shell
gem install docx
```

### Reading

``` ruby
require 'docx'

# Create a Docx::Document object for our existing docx file
doc = Docx::Document.open('example.docx')

# Retrieve and display paragraphs
doc.paragraphs.each do |p|
  puts p
end

# Retrieve and display bookmarks, returned as hash with bookmark names as keys and objects as values
doc.bookmarks.each_pair do |bookmark_name, bookmark_object|
  puts bookmark_name
end
```

Don't have a local file but a buffer? Docx handles those to:

```ruby
require 'docx'

# Create a Docx::Document object from a remote file
doc = Docx::Document.open(buffer)

# Everything about reading is the same as shown above
```

### Rendering html
``` ruby
require 'docx'

# Retrieve and display paragraphs as html
doc = Docx::Document.open('example.docx')
doc.paragraphs.each do |p|
  puts p.to_html
end
```

### Reading tables

``` ruby
require 'docx'

# Create a Docx::Document object for our existing docx file
doc = Docx::Document.open('tables.docx')

first_table = doc.tables[0]
puts first_table.row_count
puts first_table.column_count
puts first_table.rows[0].cells[0].text
puts first_table.columns[0].cells[0].text

# Iterate through tables
doc.tables.each do |table|
  table.rows.each do |row| # Row-based iteration
    row.cells.each do |cell|
      puts cell.text
    end
  end

  table.columns.each do |column| # Column-based iteration
    column.cells.each do |cell|
      puts cell.text
    end
  end
end
```

### Writing

``` ruby
require 'docx'

# Create a Docx::Document object for our existing docx file
doc = Docx::Document.open('example.docx')

# Insert a single line of text after one of our bookmarks
doc.bookmarks['example_bookmark'].insert_text_after("Hello world.")

# Insert multiple lines of text at our bookmark
doc.bookmarks['example_bookmark_2'].insert_multiple_lines_after(['Hello', 'World', 'foo'])

# Remove paragraphs
doc.paragraphs.each do |p|
  p.remove! if p.to_s =~ /TODO/
end

# Substitute text, preserving formatting
doc.paragraphs.each do |p|
  p.each_text_run do |tr|
    tr.substitute('_placeholder_', 'replacement value')
  end
end

# Save document to specified path
doc.save('example-edited.docx')
```

### Writing to tables

``` ruby
require 'docx'

# Create a Docx::Document object for our existing docx file
doc = Docx::Document.open('tables.docx')

# Iterate over each table
doc.tables.each do |table|
  last_row = table.rows.last
  
  # Copy last row and insert a new one before last row
  new_row = last_row.copy
  new_row.insert_before(last_row)

  # Substitute text in each cell of this new row
  new_row.cells.each do |cell|
    cell.paragraphs.each do |paragraph|
      paragraph.each_text_run do |text|
        text.substitute('_placeholder_', 'replacement value')
      end
    end
  end
end

doc.save('tables-edited.docx')
```

### Advanced

``` ruby
require 'docx'

d = Docx::Document.open('example.docx')

# The Nokogiri::XML::Node on which an element is based can be accessed using #node
d.paragraphs.each do |p|
  puts p.node.inspect
end

# The #xpath and #at_xpath methods are delegated to the node from the element, saving a step
p_element = d.paragraphs.first
p_children = p_element.xpath("//child::*") # selects all children
p_child = p_element.at_xpath("//child::*") # selects first child
```

## Development

### todo

* Calculate element formatting based on values present in element properties as well as properties inherited from parents
* Default formatting of inserted elements to inherited values
* Implement formattable elements.
* Implement styles.
* Easier multi-line text insertion at a single bookmark (inserting paragraph nodes after the one containing the bookmark)
