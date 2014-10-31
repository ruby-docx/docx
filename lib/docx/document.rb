require 'docx/parser'
require 'zip'

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
    delegate :paragraphs, :bookmarks, :tables, :header, :footer, :remove_unneeded_rows, :to => :@parser
    delegate :doc, :xml, :zip, :doc_header, :doc_footer, :to => :@parser
    def initialize(path, &block)
      @replace = {}
      if block_given?
        @parser = Parser.new(File.expand_path(path), &block)
      else
        @parser = Parser.new(File.expand_path(path))
      end
    end
    
    # With no associated block, Docx::Document.open is a synonym for Docx::Document.new. If the optional code block is given, it will be passed the opened +docx+ file as an argument and the Docx::Document oject will automatically be closed when the block terminates. The values of the block will be returned from Docx::Document.open.
    # call-seq:
    #   open(filepath) => file
    #   open(filepath) {|file| block } => obj
    def self.open(path, &block)
      self.new(path, &block)
    end

    ##
    # *Deprecated*
    # 
    # Iterates over paragraphs within document
    # call-seq:
    #   each_paragraph => Enumerator
    def each_paragraph
      paragraphs.each { |p| yield(p) }
    end
    
    # call-seq:
    #   to_s -> string
    def to_s
      paragraphs.map(&:to_s).join("\n")
    end

    # Save document to provided path
    # call-seq:
    #   save(filepath) => void
    def save(path)
      update
      Zip::OutputStream.open(path) do |out|
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
      @replace["word/header1.xml"] = self.header.node.first.serialize :save_with => 0 if self.header
      @replace["word/footer1.xml"] = self.footer.node.first.serialize :save_with => 0 if self.footer

      #@doc.xpath('//w:document//w:body//w:tbl//w:tr')[2].remove
    end

  end
end
