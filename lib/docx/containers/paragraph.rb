require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Paragraph
        include Container
        include Elements::Element

        def self.tag
          'p'
        end


        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        def initialize(node, document_properties = {}, doc = nil)
          @node = node
          @properties_tag = 'pPr'
          @document_properties = document_properties
          @font_size = @document_properties[:font_size]
          @document = doc
        end

        # Set text of paragraph
        def text=(content)
          if text_runs.size == 1
            text_runs.first.text = content
          elsif text_runs.size == 0
            new_r = TextRun.create_within(self)
            new_r.text = content
          else
            text_runs.each {|r| r.node.remove }
            new_r = TextRun.create_within(self)
            new_r.text = content
          end
        end

        # Return text of paragraph
        def to_s
          text_runs.map(&:text).join('')
        end

        # Return paragraph as a <p></p> HTML fragment with formatting based on properties.
        def to_html
          html = +''
          text_runs.each do |text_run|
            html << text_run.to_html
          end
          styles = { 'font-size' => "#{font_size}pt" }
          styles['color'] = "##{font_color}" if font_color
          styles['text-align'] = alignment if alignment
          html_tag(:p, content: html, styles: styles)
        end


        # Array of text runs contained within paragraph
        def text_runs
          @node.xpath('w:r|w:hyperlink').map { |r_node| Containers::TextRun.new(r_node, @document_properties) }
        end

        # Iterate over each text run within a paragraph
        def each_text_run
          text_runs.each { |tr| yield(tr) }
        end

        # Substitute text within the paragraph, even when a match spans multiple
        # text runs (e.g. a "{{placeholder}}" that Word split across several runs,
        # such as "{{fi", "rst_na", "me}}"). The per-run TextRun#substitute cannot
        # match those, but this can, because it joins the runs first.
        #
        # The substitution works on the paragraph's w:t nodes, so it also handles
        # runs holding several w:t (LibreOffice writes "text<w:br/>text" in one
        # run) and placeholders that start inside a hyperlink. The replacement
        # goes into the w:t holding the first matched character, so that run's
        # formatting is used; the rest of the match is cut from the following
        # w:t, and line breaks and tabs inside the match are dropped. Text,
        # breaks and tabs outside the match stay where they are.
        #
        # +pattern+ may be a String or a Regexp; +replacement+ follows String#sub
        # semantics, so capture-group backreferences (e.g. '\1') work with a Regexp.
        #
        #   # given a paragraph reading "Hello {{first_name}}!"
        #   paragraph.substitute('{{first_name}}', 'Jane')   # => "Hello Jane!"
        #   paragraph.substitute(/\{\{(\w+)\}\}/, 'value of \1')
        #
        # See https://github.com/ruby-docx/docx/issues/147 and #187
        def substitute(pattern, replacement)
          search_from = 0
          while (match = find_match(pattern, search_from))
            replaced = match[0].sub(pattern, replacement)
            replace_text_range(match.begin(0)...match.end(0), replaced)

            # advance past the inserted replacement so it is not re-matched
            search_from = match.begin(0) + replaced.length
          end
          self
        end

        def aligned_left?
          ['left', nil].include?(alignment)
        end

        def aligned_right?
          alignment == 'right'
        end

        def aligned_center?
          alignment == 'center'
        end

        def font_size
          size_attribute = @node.at_xpath('w:pPr//w:sz//@w:val')

          return @font_size unless size_attribute

          size_attribute.value.to_i / 2
        end

        def font_color
          color_tag = @node.xpath('w:r//w:rPr//w:color').first
          color_tag ? color_tag.attributes['val']&.value : nil
        end

        def style
          return nil unless @document

          @document.style_name_of(style_id) ||
            @document.default_paragraph_style
        end

        def style_id
          style_property.get_attribute('w:val')
        end

        def style=(identifier)
          id = @document.styles_configuration.style_of(identifier).id

          style_property.set_attribute('w:val', id)
        end

        alias_method :style_id=, :style=
        alias_method :text, :to_s

        private

        # The w:t read by #text (those of the runs and hyperlink runs), in
        # document order.
        TEXT_XPATH = 'w:r/w:t|w:hyperlink/w:r/w:t'
        # Line breaks and tabs that sit between those w:t.
        SEPARATOR_XPATH = %w[br tab cr].map { |tag| "w:r/w:#{tag}|w:hyperlink/w:r/w:#{tag}" }.join('|')
        private_constant :TEXT_XPATH, :SEPARATOR_XPATH

        # The paragraph's w:t, in document order; their contents joined make up #text.
        def text_nodes
          @node.xpath(TEXT_XPATH).map { |t_node| Elements::Text.new(t_node) }
        end

        # The first match of +pattern+ in the paragraph text at or after +from+,
        # or nil when there is none. An empty match also gives nil, so that
        # #substitute stops rather than loop on it.
        def find_match(pattern, from)
          match = text_nodes.map(&:content).join.match(pattern, from)
          match unless match.nil? || match[0].empty?
        end

        # Replace the characters in +range+ of the paragraph text with +string+.
        # The string goes into the w:t holding the range's first character; the
        # rest of the range is cut from the following w:t, and line breaks and
        # tabs inside it are removed.
        def replace_text_range(range, string)
          texts = text_nodes
          # boundaries[i]...boundaries[i + 1] is where texts[i] sits in the text
          boundaries = texts.each_with_object([0]) { |text, acc| acc << acc.last + text.content.length }
          first = texts.each_index.find { |i| boundaries[i + 1] > range.begin }
          last  = texts.each_index.find { |i| boundaries[i + 1] >= range.end }

          head = texts[first].content[0...(range.begin - boundaries[first])]
          tail = texts[last].content[(range.end - boundaries[last])..]
          if first == last
            texts[first].content = head + string + tail
            return
          end

          texts[first].content = head + string
          texts[(first + 1)...last].each { |text| text.content = '' }
          texts[last].content = tail
          remove_separators_between(texts[first].node, texts[last].node, TEXT_XPATH, SEPARATOR_XPATH)
        end

        def style_property
          properties&.at_xpath('w:pStyle') || properties&.add_child('<w:pStyle/>').first
        end

        # Returns the alignment if any, or nil if left
        def alignment
          @node.at_xpath('.//w:jc/@w:val')&.value
        end
      end
    end
  end
end
