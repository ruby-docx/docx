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
    end
  end
end
