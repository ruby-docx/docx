require 'zip/zip'

module Docx
  class Parser
    def initialize(path)
      
      yield self if block_given?
    end
    
    def paragraphs
      []
    end
  end
end
