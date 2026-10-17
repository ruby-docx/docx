require 'docx/elements'

module Docx
  module Elements
    module Containers
      module Container
        # Relation methods
        # TODO: Create a properties object, include Element
        def properties
          @node.at_xpath("./w:#{@properties_tag}")
        end

        # Erase text within an element
        def blank!
          @node.xpath('.//w:t').each { |t| t.content = '' }
        end

        def remove!
          @node.remove
        end

        private

        # Remove the line breaks and tabs (nodes matched by +separator_xpath+)
        # lying between the w:t nodes +first_t+ and +last_t+, which are matched
        # by +text_xpath+.
        def remove_separators_between(first_t, last_t, text_xpath, separator_xpath)
          separators = @node.xpath(separator_xpath).to_a
          nodes = @node.xpath("#{text_xpath}|#{separator_xpath}").to_a
          from = nodes.index(first_t)
          to = nodes.index(last_t)
          nodes[(from + 1)...to].each { |node| node.remove if separators.include?(node) }
        end
      end
    end
  end
end
