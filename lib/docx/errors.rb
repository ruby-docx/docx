module Docx
  module Errors
    StyleNotFound = Class.new(StandardError)
    StyleInvalidPropertyValue = Class.new(StandardError)
  end
end