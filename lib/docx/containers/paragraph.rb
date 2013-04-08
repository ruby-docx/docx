require 'docx/containers/text_run'
require 'docx/containers/container'

module Docx
  module Containers
    class Paragraph
      include Container
      attr_accessor :text_runs
      
      # Child elements: pPr, r, fldSimple, hlink, subDoc
      # http://msdn.microsoft.com/en-us/library/office/ee364458(v=office.11).aspx
      def initialize(node)
        @node = node
        @properties_tag = 'pPr'
        setup_text_runs
      end

      def to_s
        @text_runs.map(&:text).join('')
      end
      
      def setup_text_runs
        @text_runs ||= @node.xpath('w:r').map do |r_node|
          Containers::TextRun.new(r_node)
        end
      end

      def each_text_run
        @text_runs.each { |tr| yield(tr) }
      end
      
      alias_method :text, :to_s
    end
  end
end
