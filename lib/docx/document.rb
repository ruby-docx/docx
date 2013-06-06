require 'docx/parser'
require 'zip/zip'

module Docx
  # The Document class wraps around a docx file and provides methods to
  # interface with it.
  # 
  #   # get a Docx::Document for a docx file in the local directory
  #   doc = Docx::Document.open("test.docx")
  #   
  #   # get the text from the document
  #   puts doc.text
  #   
  #   # do the same thing in a block
  #   Docx::Document.open("test.docx") do |d|
  #     puts d.text
  #   end
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
    
    # @param [String] path
    # @return [Docx::Document]
    def self.open(path, &block)
      self.new(path, &block)
    end

    ##
    # +Deprecated+
    # 
    # Iterates over paragraphs within document
    def each_paragraph
      paragraphs.each { |p| yield(p) }
    end
    
    # @return [String]
    def to_s
      paragraphs.map(&:to_s).join("\n")
    end

    # Save document to provided path
    # call-seq:
    #   save(arg1) => void
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

    private

    #--
    # TODO: Flesh this out to be compatible with other files
    # TODO: Method to set flag on files that have been edited, probably by inserting something at the 
    # end of methods that make edits?
    #++
    def update
      @replace["word/document.xml"] = doc.serialize :save_with => 0
    end

  end
end
