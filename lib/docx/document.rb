require 'docx/parser'
require 'zip/zip'

module Docx
  class Document
    delegate :paragraphs, :bookmarks, :to => :@parser
    delegate :doc, :xml, :zip, :to => :@parser
    def initialize(path, &block)
      @replace = {}
      if block_given?
        @parser = Parser.new(File.expand_path(path), &block)
      else
        @parser = Parser.new(File.expand_path(path))
      end
    end
    
    def self.open(path, &block)
      self.new(path, &block)
    end

    def each_paragraph
      paragraphs.each { |p| yield(p) }
    end
    
    def to_s
      paragraphs.map(&:to_s).join("\n")
    end

    # TODO: Flesh this out to be compatible with other files
    # TODO: Method to set flag on files that have been edited, probably by inserting something at the 
    # end of methods that make edits?
    def update
      @replace["word/document.xml"] = doc.serialize :save_with => 0
    end

    def save(path)
      update
      Zip::ZipOutputStream.open(path) do |out|
        zip.each do |entry|
          out.put_next_entry(entry.name)

          if @replace[entry.name]
            out.write(@replace[entry.name])
          else
            out.write(zip.read(entry.name))
          end
        end
      end
      zip.close
    end
    
    alias_method :text, :to_s
  end
end
