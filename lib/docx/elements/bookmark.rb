require 'docx/elements/element'

module Docx
  module Elements
    class Bookmark
      include Element
      attr_accessor :name
      
      def initialize(node)
        @node = node
        @name = @node['w:name']
      end
    end
  end
end