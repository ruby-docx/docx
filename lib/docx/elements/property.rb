module Docx
  module Elements
    class Property
      include Element
      delegate :content, :content=, :to => :@node

      def self.tag
        'rPr'
      end


      def initialize(node)
        @node = node
      end
    end
  end
end