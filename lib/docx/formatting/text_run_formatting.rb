require 'docx/formatting/formatting'

module Docx
  module TextRunFormatting
    include Formatting

    def apply_formatting(formatting)
      if (formatting[:font])
        font_node = add_property('rFonts')
        font_node["w:ascii"] = formatting[:font]
        font_node["w:hAnsi"] = formatting[:font]
      end
      if (formatting[:font_size])
        font_size_node = add_property('sz')
        font_size_node['w:val'] = formatting[:font_size] * 2 # Font size is stored in half-points
      end
      add_property('i') if formatting[:italic]
      add_property('b') if formatting[:bold]
      add_property('u') if formatting[:underline]
      if (formatting[:color])
        color_node = add_property('color')
        color_node["w:val"] = formatting[:color]
      end
    end

    def parse_formatting()
      formatting = {}
      formatting[:italic] = !node.xpath('.//w:i').empty?
      formatting[:bold] = !node.xpath('.//w:b').empty?
      formatting[:underline] = !node.xpath('.//w:u').empty?
      font_node = node.at_xpath('.//w:rFonts')
      formatting[:font] = font_node ? font_node['w:ascii'] : document_properties[:font]
      formatting[:font_size] = font_size
      color_node = node.at_xpath('.//w:color')
      formatting[:color] = color_node ? color_node['w:val'] : nil
      formatting
    end

    def self.default_formatting(document_properties)
      {
        italic: false, bold: false, underline: false,
        font: document_properties[:font],
        font_size: document_properties[:font_size],
        color: nil
      }
    end

    def font_size
      size_tag = @node.at_xpath('.//w:sz')
      size_tag ? size_tag.attributes['val'].value.to_i / 2 : @document_properties[:font_size]
    end
  end
end
