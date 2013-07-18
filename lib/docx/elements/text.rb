module Docx
  module Elements
    class Text
      include Element
      delegate :content, :content=, :to => :@node

      def self.tag
        't'
      end


      def initialize(node)
        @node = node
      end
    end
  end
end