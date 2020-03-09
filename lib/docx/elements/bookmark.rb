require 'docx/elements/element'

module Docx
  module Elements
    class Bookmark
      include Element
      attr_accessor :name

      def self.tag
        'bookmarkStart'
      end

      def initialize(node)
        @node = node
        @name = @node['w:name']
      end

      # Insert text before bookmarkStart node
      def insert_text_before(text)
        text_run = get_run_before
        text_run.text = "#{text_run.text}#{text}"
      end

      # Insert text after bookmarkStart node
      def insert_text_after(text)
        text_run = get_run_after
        text_run.text = "#{text}#{text_run.text}"
      end

      # insert multiple lines starting with paragraph containing bookmark node.
      def insert_multiple_lines(text_array)
        # Hold paragraphs to be inserted into, corresponding to the index of the strings in the text array
        paragraphs = []
        paragraph = self.parent_paragraph
        # Remove text from paragraph
        paragraph.blank!
        paragraphs << paragraph
        for i in 0...(text_array.size - 1)
          # Copy previous paragraph
          new_p = paragraphs[i].copy
          # Insert as sibling of previous paragraph
          new_p.insert_after(paragraphs[i])
          paragraphs << new_p
        end

        # Insert text into corresponding newly created paragraphs
        paragraphs.each_index do |index|
          paragraphs[index].text = text_array[index]
        end
      end

      # Get text run immediately prior to bookmark node
      def get_run_before
        # at_xpath returns the first match found and preceding-sibling returns siblings in the
        # order they appear in the document not the order as they appear when moving out from
        # the starting node
        if not (r_nodes = @node.xpath("./preceding-sibling::w:r")).empty?
          r_node = r_nodes.last
          Containers::TextRun.new(r_node)
        else
          new_r = Containers::TextRun.create_with(self)
          new_r.insert_before(self)
          new_r
        end
      end

      # Get text run immediately after bookmark node
      def get_run_after
        if (r_node = @node.at_xpath("./following-sibling::w:r"))
          Containers::TextRun.new(r_node)
        else
          new_r = Containers::TextRun.create_with(self)
          new_r.insert_after(self)
          new_r
        end
      end
    end
  end
end