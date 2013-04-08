module Docx
  module Elements
    module Element
      def parent(type = '*')
        @node.at_xpath("./parent::#{type}")
      end

      def paragraph
        parent('w:p')
      end
    end
  end
end