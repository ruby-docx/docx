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

### Writing and Manipulating Styles
``` ruby
require 'docx'

d = Docx::Document.open('example.docx')
existing_style = d.styles_configuration.style_of("Heading 1")
existing_style.font_color = "000000"

# see attributes below
new_style = d.styles_configuration.add_style("Red", name: "Red", font_color: "FF0000", font_size: 20)
new_style.bold = true

d.paragraphs.each do |p|
  p.style = "Red"
end

d.paragraphs.each do |p|
  p.style = "Heading 1"
end

d.styles_configuration.remove_style("Red")
```

#### Style Attributes

The following is a list of attributes and what they control within the style.

- **id**: The unique identifier of the style. (required)
- **name**: The human-readable name of the style. (required)
- **type**: Indicates the type of the style (e.g., paragraph, character).
- **keep_next**: Boolean value controlling whether to keep a paragraph and the next one on the same page. Valid values: `true`/`false`.
- **keep_lines**: Boolean value specifying whether to keep all lines of a paragraph together on one page. Valid values: `true`/`false`.
- **page_break_before**: Boolean value indicating whether to insert a page break before the paragraph. Valid values: `true`/`false`.
- **widow_control**: Boolean value controlling widow and orphan lines in a paragraph. Valid values: `true`/`false`.
- **shading_style**: Defines the shading pattern style.
- **shading_color**: Specifies the color of the shading pattern. Valid values: Hex color codes.
-  **shading_fill**: Indicates the background fill color of shading.
-  **suppress_auto_hyphens**: Boolean value controlling automatic hyphenation. Valid values: `true`/`false`.
-  **bidirectional_text**: Boolean value indicating if the paragraph contains bidirectional text. Valid values: `true`/`false`.
-  **spacing_before**: Defines the spacing before a paragraph.
-  **spacing_after**: Specifies the spacing after a paragraph.
-  **line_spacing**: Indicates the line spacing of a paragraph.
-  **line_rule**: Defines how line spacing is calculated.
-  **indent_left**: Sets the left indentation of a paragraph.
-  **indent_right**: Specifies the right indentation of a paragraph.
-  **indent_first_line**: Indicates the first line indentation of a paragraph.
-  **align**: Controls the text alignment within a paragraph.
-  **font**: Sets the font for different scripts (ASCII, complex script, East Asian, etc.).
-  **font_ascii**: Specifies the font for ASCII characters.
-  **font_cs**: Indicates the font for complex script characters.
-  **font_hAnsi**: Sets the font for high ANSI characters.
-  **font_eastAsia**: Specifies the font for East Asian characters.
-  **bold**: Boolean value controlling bold formatting. Valid values: `true`/`false`.
-  **italic**: Boolean value indicating italic formatting. Valid values: `true`/`false`.
-  **caps**: Boolean value controlling capitalization. Valid values: `true`/`false`.
-  **small_caps**: Boolean value specifying small capital letters. Valid values: `true`/`false`.
-  **strike**: Boolean value indicating strikethrough formatting. Valid values: `true`/`false`.
-  **double_strike**: Boolean value defining double strikethrough formatting. Valid values: `true`/`false`.
-  **outline**: Boolean value specifying outline effects. Valid values: `true`/`false`.
-  **outline_level**: Indicates the outline level in a document's hierarchy.
-  **font_color**: Sets the text color. Valid values: Hex color codes.
-  **font_size**: Controls the font size.
-  **font_size_cs**: Specifies the font size for complex script characters.
-  **underline_style**: Indicates the style of underlining.
-  **underline_color**: Specifies the color of the underline. Valid values: Hex color codes.
-  **spacing**: Controls character spacing.
-  **kerning**: Sets the space between characters.
-  **position**: Controls the position of characters (superscript/subscript).
-  **text_fill_color**: Sets the fill color of text. Valid values: Hex color codes.
-  **vertical_alignment**: Controls the vertical alignment of text within a line.
-  **lang**: Specifies the language tag for the text.

## Development

### todo

* Calculate element formatting based on values present in element properties as well as properties inherited from parents
* Default formatting of inserted elements to inherited values
* Implement formattable elements.
* Easier multi-line text insertion at a single bookmark (inserting paragraph nodes after the one containing the bookmark)
