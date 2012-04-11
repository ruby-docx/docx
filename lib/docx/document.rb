require 'docx/parser'

module Docx
  class Document
    attr_reader :paragraphs
    
    def initialize(path)
      Parser.new(File.expand_path(path)) do |p|
        @paragraphs = p.paragraphs
      end
    end
    
    def self.open(path)
      self.new(path)
    end
    
    def each_paragraph
      @paragraphs.each { |p| yield(p) }
    end
    
    def to_s
      @paragraphs.map(&:to_s).join("\n")
    end
    
    alias_method :text, :to_s
  end
end
