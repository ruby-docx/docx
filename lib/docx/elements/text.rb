module Docx
  module Elements
    class Text
      include Element
      delegate :content, :content=, :to => :@node
      TAG = 't'

      def initialize(node)
        @node = node
      end
    end
  end
end