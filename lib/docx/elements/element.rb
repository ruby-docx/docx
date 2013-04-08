require 'nokogiri'

module Docx
  module Elements
    module Element
      def self.included(base)
        base.extend(ClassMethods)
      end

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
      def copy
        self.class.new(@node.dup)
      end

      module ClassMethods
        def create_with(element)
          # Need to somehow get the xml document accessible here by default, but this is alright in the interim
          self.new(Nokogiri::XML::Node.new("w:#{TAG}"), element.node)
        end

        def create_within(element)
          new_element = create_with(element)
          new_element.append_to(element)
        end
      end
    end
  end
end