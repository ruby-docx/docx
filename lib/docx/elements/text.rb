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
        # Without xml:space="preserve", Word drops leading and trailing spaces.
        @node['xml:space'] = 'preserve' if args.to_s.match?(/\A\s|\s\z/)
      end

      def initialize(node)
        @node = node
      end
    end
  end
end
