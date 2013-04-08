require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class Paragraph
        include Container
        attr_accessor :text_runs
        
        TAG = 'p'

        # Child elements: pPr, r, fldSimple, hlink, subDoc
        # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
        def initialize(node)
          @node = node
          @properties_tag = 'pPr'
        end

        # Handle direct text insertion into paragraph on some conditions
        def text=(content)
          if @text_runs.size == 1
            @text_runs.first.text = content
          elsif @text_runs.size == 0
            new_r = TextRun.create_within(self)
            new_r.text = content
          else
            nil
          end
        end

        def to_s
          text_runs.map(&:text).join('')
        end

        def text_runs
          @node.xpath('w:r').map {|r_node| Containers::TextRun.new(r_node) }
        end

        def each_text_run
          text_runs.each { |tr| yield(tr) }
        end
        
        alias_method :text, :to_s
      end
    end
  end
end
