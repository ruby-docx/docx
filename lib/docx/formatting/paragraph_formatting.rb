require 'docx/formatting/formatting'

module Docx
  module ParagraphFormatting
    include Formatting

    def apply_formatting(formatting)
      if (formatting[:alignment])
        alignment_node = add_property('jc')
        alignment_node['w:val'] = formatting[:alignment]
      end
    end

    def parse_formatting
      formatting = {}
      alignment_node = node.at_xpath('.//w:jc')
      formatting[:alignment] = alignment_node ? alignment_node['w:val'] : nil
      formatting
    end

    def self.default_formatting
      {
        alignment: nil
      }
    end
  end
end
