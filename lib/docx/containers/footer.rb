require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Footer
        include Container
        include Elements::Element

        def self.tag
          'ftr'
        end

        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        def initialize(node)
          @node = node
          @properties_tag = 'pPr'
        end

        def to_s
          paragraphs.map(&:text).join('')
        end

        def paragraphs
          @node.xpath('w:p').map {|p_node| Containers::Paragraph.new(p_node) }
        end

        def each_paragraph
          paragraphs.each { |tr| yield(tr) }
        end
        
        alias_method :text, :to_s
      end
    end
  end
end
