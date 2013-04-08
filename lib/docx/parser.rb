require 'docx/containers'
require 'docx/elements'
require 'nokogiri'
require 'zip/zip'

module Docx
  class Parser
    attr_reader :xml, :doc, :zip
    def initialize(path)
      @zip = Zip::ZipFile.open(path)
      @xml = @zip.read('word/document.xml')
      @doc = Nokogiri::XML(@xml)
      if block_given?
        yield self
        @zip.close
      end
    end
    
    def paragraphs
      @doc.xpath('//w:document//w:body//w:p').map { |p_node| parse_paragraph_from p_node }
    end

    def bookmarks
      @doc.xpath('//w:bookmarkStart').map { |b_node| parse_bookmark_from b_node }
    end
    
    private
    
    def parse_paragraph_from(p_node)
      Elements::Containers::Paragraph.new(p_node)
    end

    def parse_bookmark_from(b_node)
      Elements::Bookmark.new(b_node)
    end
  end
end
