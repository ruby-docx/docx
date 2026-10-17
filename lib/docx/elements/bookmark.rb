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
        append_to_run(get_run_before, text)
      end

      # Insert text after bookmarkStart node
      def insert_text_after(text)
        prepend_to_run(get_run_after, text)
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
        if enclosing_paragraph
          # at_xpath returns the first match found and preceding-sibling returns siblings in the
          # order they appear in the document not the order as they appear when moving out from
          # the starting node
          unless (r_nodes = @node.xpath("./preceding-sibling::w:r")).empty?
            return Containers::TextRun.new(r_nodes.last)
          end

          new_r = Containers::TextRun.create_with(self)
          new_r.insert_before(self)
          return new_r
        end

        # Block-level bookmark (e.g. Google Docs places bookmarkStart/End directly
        # under w:body). Add a run to the preceding paragraph, or to a new
        # paragraph before the bookmark, so the run lives inside a w:p.
        run_at_paragraph_end(@node.xpath("./preceding-sibling::w:p").last) ||
          run_in_new_paragraph { |paragraph| @node.add_previous_sibling(paragraph) }
      end

      # Get text run immediately after bookmark node
      def get_run_after
        if enclosing_paragraph
          if (r_node = @node.at_xpath("./following-sibling::w:r"))
            return Containers::TextRun.new(r_node)
          end

          new_r = Containers::TextRun.create_with(self)
          new_r.insert_after(self)
          return new_r
        end

        # Block-level bookmark: add a run to the following paragraph, or to a new
        # paragraph after the bookmark.
        run_at_paragraph_start(@node.at_xpath("./following-sibling::w:p")) ||
          run_in_new_paragraph { |paragraph| @node.add_next_sibling(paragraph) }
      end

      # Override Element#parent_paragraph so insert_multiple_lines also works for
      # block-level bookmarks. For those we fill a fresh paragraph inserted at the
      # bookmark, rather than an adjacent existing paragraph, so we neither crash
      # (no enclosing paragraph) nor overwrite unrelated content.
      def parent_paragraph
        return Containers::Paragraph.new(enclosing_paragraph) if enclosing_paragraph

        paragraph = Nokogiri::XML::Node.new("w:p", @node.document)
        @node.add_next_sibling(paragraph)
        Containers::Paragraph.new(paragraph)
      end

      private

      # A run's text, line breaks and tabs, in document order.
      RUN_CONTENT_XPATH = 'w:t|w:br|w:tab|w:cr'
      private_constant :RUN_CONTENT_XPATH

      # Add text at the very end of the run, which sits just before the
      # bookmark: to its last w:t, or to a new w:t when the run ends with a
      # break or tab. The rest of the run is left as it is (#197).
      def append_to_run(run, text)
        last = run.node.xpath(RUN_CONTENT_XPATH).last
        if last && run.node.xpath('w:t').include?(last)
          Text.new(last).content = last.content + text
        else
          Text.create_within(run).content = text
        end
      end

      # Add text at the very start of the run, which sits just after the
      # bookmark: to its first w:t, or to a new w:t when the run starts with
      # a break or tab. The rest of the run is left as it is (#197).
      def prepend_to_run(run, text)
        first = run.node.xpath(RUN_CONTENT_XPATH).first
        if first && run.node.xpath('w:t').include?(first)
          Text.new(first).content = text + first.content
        elsif first
          new_t = Text.create_with(run)
          new_t.node = first.add_previous_sibling(new_t.node)
          new_t.content = text
        else
          Text.create_within(run).content = text
        end
      end

      # The w:p the bookmark sits inside, or nil when it is block-level.
      def enclosing_paragraph
        @node.at_xpath("./parent::w:p")
      end

      # A new run inserted at the start of paragraph_node (after w:pPr if present).
      # Returns nil when paragraph_node is nil.
      def run_at_paragraph_start(paragraph_node)
        return nil unless paragraph_node

        new_r = Containers::TextRun.create_with(self)
        if (props = paragraph_node.at_xpath("w:pPr"))
          props.add_next_sibling(new_r.node)
        else
          paragraph_node.prepend_child(new_r.node)
        end
        new_r
      end

      # A new run appended to the end of paragraph_node. nil when it is nil.
      def run_at_paragraph_end(paragraph_node)
        return nil unless paragraph_node

        Containers::TextRun.create_within(Containers::Paragraph.new(paragraph_node))
      end

      # Create a run wrapped in a fresh w:p; the block positions that paragraph
      # relative to the bookmark.
      def run_in_new_paragraph
        new_r = Containers::TextRun.create_with(self)
        paragraph = Nokogiri::XML::Node.new("w:p", @node.document)
        paragraph.add_child(new_r.node)
        yield paragraph
        new_r
      end
    end
  end
end