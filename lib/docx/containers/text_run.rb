require 'docx/containers/container'
require 'docx/formatting/text_run_formatting'

module Docx
  module Elements
    module Containers
      class TextRun
        include Container
        include Elements::Element
        include TextRunFormatting

        def self.tag
          'r'
        end

        attr_reader :text
        attr_reader :document_properties
        attr_reader :properties_tag
        alias_method :formatting, :parse_formatting

        def initialize(node, document_properties = {})
          @node = node
          @document_properties = document_properties
          @text_nodes = @node.xpath('w:t').map {|t_node| Elements::Text.new(t_node) }
          @properties_tag = 'rPr'
          @text       = parse_text || ''
          @font_size = @document_properties[:font_size]
        end

        # Set text of text run
        def text=(content)
          if @text_nodes.size == 1
            @text_nodes.first.content = content
          elsif @text_nodes.empty?
            new_t = Elements::Text.create_within(self)
            new_t.content = content
          end
        end

        # Set the text of text run with formatting
        def set_text(content, formatting={})
          self.text = content
          apply_formatting(formatting)
        end

        # Returns text contained within text run
        def parse_text
          @text_nodes.map(&:content).join('')
        end

        def to_s
          @text
        end

        # Return text as a HTML fragment with formatting based on properties.
        def to_html
          html = @text
          html = html_tag(:em, content: html) if italicized?
          html = html_tag(:strong, content: html) if bolded?
          styles = {}
          styles['text-decoration'] = 'underline' if underlined?
          # No need to be granular with font size down to the span level if it doesn't vary.
          styles['font-size'] = "#{font_size}pt" if font_size != @font_size
          styles['font-family'] = %Q["#{formatting[:font]}"] if formatting[:font]
          styles['color'] = "##{formatting[:color]}" if formatting[:color]
          html = html_tag(:span, content: html, styles: styles) unless styles.empty?
          return html
        end

        def italicized?
          formatting[:italic]
        end

        def bolded?
          formatting[:bold]
        end

        def underlined?
          formatting[:underline]
        end
      end
    end
  end
end
