require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class TableCell
        include Container
        include Elements::Element

        def self.tag
          'tc'
        end

        def initialize(node)
          @node = node
          @properties_tag = 'tcPr'
        end

        # Return text of paragraph's cell
        def to_s
          paragraphs.map(&:text).join('')
        end

        # Array of paragraphs contained within cell
        def paragraphs
          @node.xpath('w:p').map {|p_node| Containers::Paragraph.new(p_node) }
        end

        # Iterate over each text run within a paragraph's cell
        def each_paragraph
          paragraphs.each { |tr| yield(tr) }
        end
        
        alias_method :text, :to_s
      end
    end
  end
end
