require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Paragraph
        include Container
        include Elements::Element

        PLACEHOLDER_REGEX = /\{\{(.*?)\}\}/ # In order to combine text runs with {{}} pattern

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
          placeholder_position_hash = detect_placeholder_positions
          content_size = [0]
          text_runs.each_with_index do |text_node, index|
            content_size[index + 1] = text_node.text.length + (index.zero? ? 0 : content_size[index])
          end
          content_size.pop
          placeholder_position_hash.each do |placeholder, placeholder_positions|
            placeholder_positions.each do |p_start_index|
              p_end_index = (p_start_index + placeholder.length - 1)
              tn_start_index = content_size.index(content_size.select { |size| size <= p_start_index }.max)
              tn_end_index = content_size.index(content_size.select { |size| size <= p_end_index }.max)
              next if tn_start_index == tn_end_index
              replace_incorrect_placeholder_content(placeholder, tn_start_index, tn_end_index, content_size[tn_start_index] - p_start_index, p_end_index - content_size[tn_end_index])
            end
          end
        end

        # This method detect the placeholder's starting index and return the starting index in array.
        # Ex: Assumptions : text = 'This is Placeholder Text with {{Placeholder}} {{Text}} {{Placeholder}}'
        #     It will detect the placeholder's starting index from the given text.
        #     Here, starting index of '{{Placeholder}}' => [30, 55], '{{Text}}' => [46]
        # @return [Hash]
        # Ex: {'{{Placeholder}}' => [30, 55], '{{Text}}' => [46]}
        def detect_placeholder_positions
          text.scan(PLACEHOLDER_REGEX).flatten.uniq.each_with_object({}) do |placeholder, placeholder_hash|
            next if placeholder.include?("{") || placeholder.include?("}")
            placeholder_text = "{{#{placeholder}}}"
            current_index = text.index(placeholder_text)
            arr_of_index = [current_index]
            until current_index.nil?
              current_index = text.index(placeholder_text, current_index + 1)
              arr_of_index << current_index unless current_index.nil?
            end
            placeholder_hash[placeholder_text] = arr_of_index
          end
        end

        # @param [String] :placeholder
        # @param [Integer] :start_index, end_index, p_start_index, p_end_index
        # This Method replaces below :
        #   1. Corrupted text nodes content with empty string
        #   2. Proper Placeholder content within the same text node
        # Ex: Assume we have a array of text nodes content as text_runs = ['This is ', 'Placeh', 'older Text', 'with ', '{{', 'Place', 'holder}}' , '{{Text}}', '{{Placeholder}}']
        #   Here if you see, the '{{placeholder}}' is not available in the same text node. We need to merge the content of indexes - text_runs[5], text_runs[6], text_runs[7].
        #   So We will replace the content as below:
        #     1. text_runs[5] = '{{Placeholder}}'
        #     2. text_runs[6] = ''
        #     3. text_runs[7] = ''
        def replace_incorrect_placeholder_content(placeholder, start_index, end_index, p_start_index, p_end_index)
          (start_index..end_index).each do |index|
            if index == start_index
              current_text = text_runs[index].text.to_s
              current_text[p_start_index..-1] = placeholder
              text_runs[index].text = current_text
            elsif index == end_index
              current_text = text_runs[index].text.to_s
              current_text[0..p_end_index] = ""
              text_runs[index].text = current_text
            else
              text_runs[index].text = ""
            end
          end
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
