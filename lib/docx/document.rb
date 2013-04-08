require 'docx/parser'

module Docx
  class Document
    #delegate :paragraphs, :bookmarks, :to => :@parser

    def initialize(path, &block)
      if block_given?
        @parser = Parser.new(File.expand_path(path), &block)
      else
        @parser = Parser.new(File.expand_path(path))
      end
    end
    
    def self.open(path, &block)
      self.new(path, &block)
    end
    
    def xml
      @parser.xml
    end

    def bookmarks
      @parser.bookmarks
    end

    def paragraphs
      @parser.paragraphs
    end

    def each_paragraph
      paragraphs.each { |p| yield(p) }
    end
    
    def to_s
      paragraphs.map(&:to_s).join("\n")
    end
    
    alias_method :text, :to_s
  end
end
