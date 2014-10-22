require 'docx/elements'

module Docx
  module Elements
    module Containers
      module Container
        # Relation methods
        # TODO: Create a properties object, include Element
        def properties
          @node.at_xpath("./#{@properties_tag}")
        end

        # Erase text within an element
        def blank!
          @node.xpath(".//w:t").each {|t| t.content = '' }
        end

        def remove!
          @node.remove
        end
      end
    end
  end
end
