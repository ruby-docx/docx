module Docx
  module Containers
    module Container
      def properties
        @node.at_xpath("./#{@properties_tag}")
      end

      def parent(type = '*')
        @node.at_xpath("./parent::#{type}")
      end

      def copy
        @node.dup
      end
    end
  end
end