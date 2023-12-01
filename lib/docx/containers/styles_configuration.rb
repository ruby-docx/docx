require 'docx/containers/container'
require 'docx/elements/style'

module Docx
  module Elements
    module Containers
      StyleNotFound = Class.new(StandardError)

      class StylesConfiguration
        def initialize(raw_styles)
          @raw_styles = raw_styles
          @styles_parent_node = raw_styles.root
        end

        attr_reader :styles, :styles_parent_node

        def styles
          styles_parent_node
            .children
            .filter_map do |style|
              next unless style.get_attribute("w:styleId")

              Elements::Style.new(self, style)
            end
        end

        def style_of(id_or_name)
          styles.find { |style| style.id == id_or_name || style.name == id_or_name } || raise(Errors::StyleNotFound, "Style name or id '#{id_or_name}' not found")
        end

        def size
          styles.size
        end

        def add_style(id, attributes = {})
          Elements::Style.create(self, {id: id, name: id}.merge(attributes))
        end

        def remove_style(id)
          style = styles.find { |style| style.id == id }

          style.node.remove
          styles.delete(style)
        end

        def serialize(**options)
          @raw_styles.serialize(**options)
        end
      end
    end
  end
end