module Docx
  module Elements
    class Style
      include Element
      delegate :content, :content=, :to => :@node

      def self.tag
        'rFonts'
      end

      def initialize(node)
        @node = node
      end
    end
  end
end