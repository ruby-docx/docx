module Docx
  module Elements
    class Text
      include Element

      def self.tag
        't'
      end

      def content
        @node.content
      end

      def content=(args)
        @node.content = args
      end

      def initialize(node)
        @node = node
      end
    end
  end
end
