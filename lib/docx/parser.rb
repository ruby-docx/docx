require 'nokogiri'
require 'zip/zip'

module Docx
  class Parser
    def initialize(path)
      @zip = Zip::ZipFile.open(path)
      @xml = Nokogiri::XML(@zip.find_entry('word/document.xml').get_input_stream)
      
      if block_given?
        yield self
        @zip.close
      end
    end
    
    def paragraphs
      @xml.xpath('//w:p').map(&:text)
    end
  end
end
