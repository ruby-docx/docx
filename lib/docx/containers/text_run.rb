module Docx
  module Containers
    class TextRun
      attr_reader :text
      attr_reader :formatting
      
      def initialize(attrs)
        @text       = attrs[:text] || ''
        @formatting = attrs[:formatting] || :none
      end
      
      def to_s
        @text
      end
    end
  end
end
