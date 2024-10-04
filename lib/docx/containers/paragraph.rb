require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Paragraph
        include Container
        include Elements::Element

        PLACEHOLDER_REGEX = /\{\{([^{}]*?)\}\}/

        def self.tag
          'p'
        end


        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        #
        # See: https://github.com/ruby-docx/docx/issues/147 for placeholder patching
        def initialize(node, document_properties = {}, doc = nil)
          @node = node
          @properties_tag = 'pPr'
          @document_properties = document_properties
          @font_size = @document_properties[:font_size]
          @document = doc
          validate_placeholder_content
        end

        def validate_placeholder_content
          # First, build a map of all text run contents and their positions
          content_map = build_content_map
          full_text = text_runs.map(&:text).join('')

          # Use global regex to find all placeholders with their positions
          placeholders = full_text.to_enum(:scan, PLACEHOLDER_REGEX).map do
            [Regexp.last_match.begin(0), Regexp.last_match.end(0)]
          end

          placeholders.each do |start_pos, end_pos|
            # Find the indexes of the text runs that includes the start and end of the placeholder
            start_text_run_index = content_map.index { |m| m[:start] <= start_pos && m[:end] >= start_pos }
            end_text_run_index = content_map.index { |m| m[:start] <= end_pos - 1 && m[:end] >= end_pos - 1 }

            next if start_text_run_index.nil? || end_text_run_index.nil?
            next if start_text_run_index == end_text_run_index # Skip if entire placeholder is already in single run

            placeholder_content = full_text[start_pos...end_pos]

            (start_text_run_index..end_text_run_index).each do |i|
              if i == start_text_run_index
                # Merge the entire placeholder into the first run
                current_text = content_map[i][:text].dup
                current_text[start_pos - content_map[i][:start]..-1] = placeholder_content
                content_map[i][:run].text = current_text
              elsif i == end_text_run_index
                # Last run should preserve any content after the placeholder
                current_text = content_map[i][:text].dup
                remaining_text = current_text[(end_pos) - content_map[i][:start]..-1]
                content_map[i][:run].text = remaining_text
              else
                # Clear intermediate runs
                content_map[i][:run].text = ''
              end
            end
          end
        end

        def build_content_map
          content_map = []
          current_position = 0

          text_runs.each do |text_run|
            run_text = text_run.text.to_s
            content_map << {
              start: current_position,
              end: current_position + run_text.length - 1,
              text: run_text,
              run: text_run
            }
            current_position += run_text.length
          end
          content_map
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
          html = ''
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
          color_tag ? color_tag.attributes['val'].value : nil
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
