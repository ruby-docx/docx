require 'docx/containers/container'

module Docx
  module Containers
    class TextRun
      include Container
      DEFAULT_FORMATTING = {
        italic:    false,
        bold:      false,
        underline: false
      }
      
      attr_reader :text
      attr_reader :formatting
      
      def initialize(node)
        @node = node
        @properties_tag = 'rPr'
        @text       = parse_text || ''
        @formatting = parse_formatting || DEFAULT_FORMATTING
      end

      def parse_text
        @node.xpath('w:t').map(&:text).join('')
      end

      def parse_formatting
        {
          italic:    !@node.xpath('.//w:i').empty?,
          bold:      !@node.xpath('.//w:b').empty?,
          underline: !@node.xpath('.//w:u').empty?
        }
      end

      def to_s
        @text
      end
      
      def italicized?
        @formatting[:italic]
      end
      
      def bolded?
        @formatting[:bold]
      end
      
      def underlined?
        @formatting[:underline]
      end
    end
  end
end
