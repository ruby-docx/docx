require 'zip/zip'

module Docx
  class Parser
    def initialize(path)
      @zip = Zip::ZipFile.open(path)
      
      if block_given?
        yield self
        @zip.close
      end
    end
    
    def paragraphs
      []
    end
  end
end
