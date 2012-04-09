module Docx
  class Document
    attr_reader :paragraphs
    
    def initialize(path)
      @paragraphs = []
    end
    
    def self.open(path)
      self.new(path)
    end
    
    def each_paragraph
      @paragraphs.each { |p| yield(p) }
    end
  end
end
