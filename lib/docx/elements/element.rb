module Docx
  module Elements
    module Element
      attr_accessor :node
      delegate :at_xpath, :xpath, :to => :@node

      # TODO: Should create a docx object from this
      def parent(type = '*')
        @node.at_xpath("./parent::#{type}")
      end

      # TODO: Should create a docx paragraph from this
      def paragraph
        parent('w:p')
      end

      # Insertion methods
      # Insert node as last child
      def append_to(element)
        @node = element.node.add_child(@node)
      end

      # Insert node as first child (after properties)
      def prepend_to(element)
        @node = element.node.properties.add_next_sibling(@node)
      end

      def insert_after(element)
        # Returns newly re-parented node
        @node = element.node.add_next_sibling(@node)
      end

      def insert_before(element)
        @node = element.node.add_previous_sibling(@node)
      end

      # Creation/edit methods
      # TODO: Need to figure out whether to return a copied node or an instance of the class. I'm thinking an instance of the class

      def copy
        self.class.new(@node.dup)
      end
    end
  end
end