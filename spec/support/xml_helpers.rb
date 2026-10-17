# frozen_string_literal: true

require 'nokogiri'

# Builds Nokogiri nodes from WordprocessingML fragments, so element and
# container classes can be unit-tested without opening a .docx file.
#
#   paragraph_node = w_node('<w:p><w:r><w:t>hello</w:t></w:r></w:p>')
#   Docx::Elements::Containers::Paragraph.new(paragraph_node)
module XmlHelpers
  NAMESPACES = {
    'w' => 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r' => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
  }.freeze

  # Wraps +xml+ in a w:document/w:body declaring the usual namespaces and
  # returns the w:body node.
  def w_body(xml)
    namespaces = NAMESPACES.map { |prefix, uri| %(xmlns:#{prefix}="#{uri}") }.join(' ')
    Nokogiri::XML(%(<w:document #{namespaces}><w:body>#{xml}</w:body></w:document>)).at_xpath('//w:body')
  end

  # The first element of +xml+, parsed inside a w:body (see #w_body).
  def w_node(xml)
    w_body(xml).element_children.first
  end
end

RSpec.configure do |config|
  config.include XmlHelpers
end
