require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Paragraph
        include Container
        include Elements::Element

        def self.tag
          'p'
        end


        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        def initialize(node)
          @node = node
          @properties_tag = 'pPr'
        end

        # Set text of paragraph
        def text=(content)
          if text_runs.size == 1
            text_runs.first.text = content
          elsif text_runs.size == 0
            new_r = TextRun.create_within(self)
            new_r.text = content
          else
            text_runs.each {|r| r.node.remove }
            new_r = TextRun.create_within(self)
            new_r.text = content
          end
        end

        # Return text of paragraph
        def to_s
          text_runs.map(&:text).join('')
        end

        # Array of text runs contained within paragraph
        def text_runs
          @node.xpath('w:r|w:hyperlink/w:r').map {|r_node| Containers::TextRun.new(r_node) }
        end

        # Iterate over each text run within a paragraph
        def each_text_run
          text_runs.each { |tr| yield(tr) }
        end
        
        alias_method :text, :to_s
      end
    end
  end
end
