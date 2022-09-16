require 'docx/containers'
require 'docx/elements'
require 'nokogiri'
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

    # A path with * indicates that there are possibly multiple documents
    # matching that glob, eg. word/header1.xml, word/header2.xml
    DOCUMENT_PATHS = {
      doc: "word/document.xml",
      styles: "word/styles.xml",
      headers: "word/header*.xml",
      footers: "word/footer*.xml",
      numbering: "word/numbering.xml",
      settings: "word/settings.xml"
    }

    attr_reader :xml, :doc, :zip, :styles, :headers, :footers, :numbering, :settings

    def initialize(path, &block)
      @replace = {}
      @zip = Zip::File.open(path)
      extract_documents

      if block_given?
        yield self
        @zip.close
      end
    end


    # This stores the current global document properties, for now
    def document_properties
      {
        font_size: font_size
      }
    end


    # With no associated block, Docx::Document.open is a synonym for Docx::Document.new. If the optional code block is given, it will be passed the opened +docx+ file as an argument and the Docx::Document oject will automatically be closed when the block terminates. The values of the block will be returned from Docx::Document.open.
    # call-seq:
    #   open(filepath) => file
    #   open(filepath) {|file| block } => obj
    def self.open(path, &block)
      self.new(path, &block)
    end

    def paragraphs
      @doc.xpath('//w:document//w:body//w:p').map { |p_node| parse_paragraph_from p_node }
    end

    def bookmarks
      bkmrks_hsh = Hash.new
      bkmrks_ary = @doc.xpath('//w:bookmarkStart').map { |b_node| parse_bookmark_from b_node }
      bkmrks_ary += @headers.values.map { |xml_doc| xml_doc.xpath('//w:bookmarkStart').map { |b_node| parse_bookmark_from b_node } }.flatten
      bkmrks_ary += @footers.values.map { |xml_doc| xml_doc.xpath('//w:bookmarkStart').map { |b_node| parse_bookmark_from b_node } }.flatten
      # auto-generated by office 2010
      bkmrks_ary.reject! {|b| b.name == "_GoBack" }
      bkmrks_ary.each {|b| bkmrks_hsh[b.name] = b }
      bkmrks_hsh
    end

    def tables
      @doc.xpath('//w:document//w:body//w:tbl').map { |t_node| parse_table_from t_node }
    end

    # Some documents have this set, others don't.
    # Values are returned as half-points, so to get points, that's why it's divided by 2.
    def font_size
      size_tag = @styles.xpath('//w:docDefaults//w:rPrDefault//w:rPr//w:sz').first
      size_tag ? size_tag.attributes['val'].value.to_i / 2 : nil
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

    # Output entire document as a String HTML fragment
    def to_html
      paragraphs.map(&:to_html).join('\n')
    end

    # Save document to provided path
    # call-seq:
    #   save(filepath) => void
    def save(path)
      update
      Zip::OutputStream.open(path) do |out|
        zip.each do |entry|
          next unless entry.file?
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

    def replace_entry(entry_path, file_contents)
      @replace[entry_path] = file_contents
    end

    private

    def extract_documents
      DOCUMENT_PATHS.each do |attr_name, path|
        if path.match /\*/
          extract_multiple_documents_from_globbed_path(attr_name, path)
        else
          extract_single_document_from_path(attr_name, path)
        end
      end
    end

    def extract_single_document_from_path(attr_name, path)
      if @zip.find_entry(path)
        xml_doc = @zip.read(path)
        self.instance_variable_set(:"@#{attr_name}", Nokogiri::XML(xml_doc))
      end
    end

    def extract_multiple_documents_from_globbed_path(hash_attr_name, glob_path)
      files = @zip.glob(glob_path).map { |h| h.name }
      filename_and_contents_pairs = files.map do |file|
        simple_file_name = file.sub(/^word\//, "").sub(/\.xml$/, "")
        [simple_file_name, Nokogiri::XML(@zip.read(file))]
      end
      hash = Hash[filename_and_contents_pairs]
      self.instance_variable_set(:"@#{hash_attr_name}", hash)
    end

    #--
    # TODO: Flesh this out to be compatible with other files
    # TODO: Method to set flag on files that have been edited, probably by inserting something at the
    # end of methods that make edits?
    #++
    def update
      DOCUMENT_PATHS.each do |attr_name, path|
        if path.match /\*/
          self.instance_variable_get("@#{attr_name}").each do |simple_file_name, contents|
            replace_entry("word/#{simple_file_name}.xml", contents.serialize(:save_with => 0))
          end
        else
          xml_document = self.instance_variable_get("@#{attr_name}")
          replace_entry path, xml_document.serialize(:save_with => 0) if xml_document
        end
      end
    end

    # generate Elements::Containers::Paragraph from paragraph XML node
    def parse_paragraph_from(p_node)
      Elements::Containers::Paragraph.new(p_node, document_properties)
    end

    # generate Elements::Bookmark from bookmark XML node
    def parse_bookmark_from(b_node)
      Elements::Bookmark.new(b_node)
    end

    def parse_table_from(t_node)
      Elements::Containers::Table.new(t_node)
    end
  end
end
