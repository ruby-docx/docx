require 'docx/containers/text_run'

module Docx
  module Containers
    class Paragraph
      attr_accessor :text_runs
      
      def initialize(txt_runs)
        @text_runs = txt_runs
      end
      
      def to_s
        @text_runs.map(&:text).join('')
      end
      
      def each_text_run
        @text_runs.each { |tr| yield(tr) }
      end
      
      alias_method :text, :to_s
    end
  end
end
