module Docx
  module Containers
    class TextRun
      DEFAULT_FORMATTING = {
        italic:    false,
        bold:      false,
        underline: false
      }
      
      attr_reader :text
      attr_reader :formatting
      
      def initialize(attrs)
        @text       = attrs[:text] || ''
        @formatting = attrs[:formatting] || DEFAULT_FORMATTING
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
