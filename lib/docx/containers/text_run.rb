require 'nokogiri'
require 'docx/elements'
require 'docx/containers/container'

module Docx
  module Elements
    module Containers
      class TextRun
        include Container
        include Elements::Element

        DEFAULT_FORMATTING = {
          italic:    false,
          bold:      false,
          underline: false
        }
        
        def self.tag
          'r'
        end

        attr_reader :text
        attr_reader :formatting
        
        def initialize(node)
          @node = node
          @text_nodes = @node.xpath('w:t').map {|t_node| Elements::Text.new(t_node) }
          @properties_tag = 'rPr'
          @text       = parse_text || ''
          @formatting = parse_formatting || DEFAULT_FORMATTING
        end

        # Set text of text run
        def text=(content)
          #binding.pry if content.include?("Max")
          if @text_nodes.size == 1
            @text_nodes.first.content = content
          elsif @text_nodes.empty?
            new_t = Elements::Text.create_within(self)
            new_t.content = content

            # attr_node = self.node.parent.children.xpath('w:rPr').map(&:children).flatten.first
            # if attr_node
            #   font = attr_node.attribute('hAnsi').value
            #   p = Elements::Property.create_within(self)
            #   st = Elements::Style.create_within(p)
            #   #binding.pry if content.include?("Max")
            #   st.node.set_attribute('w:ascii', font)
            #   st.node.set_attribute('w:hAnsi', font)
            # end
          end
        end

        # Returns text contained within text run
        def parse_text
          @text_nodes.map(&:content).join('')
        end

        def parse_formatting
          {
            italic:    !@node.xpath('.//w:i').empty?,
            bold:      !@node.xpath('.//w:b').empty?,
            underline: !@node.xpath('.//w:u').empty?
          }
        end

        def to_s
          @text
        end
        
        def italicized?
          @formatting[:italic]
        end
        
        def bolded?
          @formatting[:bold]
        end
        
        def underlined?
          @formatting[:underline]
        end
      end
    end
  end
end
