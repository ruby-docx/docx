require 'docx/helpers'
require 'docx/elements'
require 'docx/elements/style/converters'
require 'docx/elements/style/validators'

module Docx
  module Elements
    class Style
      include Docx::SimpleInspect

      class Attribute
        attr_reader :name, :selectors, :required, :converter, :validator

        def initialize(name, selectors, required: false, converter:, validator:)
          @name = name
          @selectors = selectors
          @required = required
          @converter = converter || Converters::DefaultValueConverter
          @validator = validator || Validators::DefaultValidator
        end

        def required?
          required
        end

        def retrieve_from(style)
          selectors
            .lazy
            .filter_map { |node_xpath| style.node.at_xpath(node_xpath)&.value }
            .map { |value| converter.decode(value) }
            .first
        end

        def assign_to(style, value)
          (required && value.nil?) &&
            raise(Errors::StyleRequiredPropertyValue, "Required value #{name}")

          validator.validate(value) ||
            raise(Errors::StyleInvalidPropertyValue, "Invalid value for #{name}: '#{value.nil? ? "nil" : value}'")

          encoded_value = converter.encode(value)

          selectors.map do |attribute_xpath|
            if (existing_attribute = style.node.at_xpath(attribute_xpath))
              if encoded_value.nil?
                existing_attribute.remove
              else
                existing_attribute.value = encoded_value.to_s
              end

              next encoded_value
            end

            next encoded_value if encoded_value.nil?

            node_xpath, attribute = attribute_xpath.split("/@")

            created_node =
              node_xpath
                .split("/")
                .reduce(style.node) do |parent_node, child_xpath|
                  # find the child node
                  parent_node.at_xpath(child_xpath) ||
                    # or create the child node
                    Nokogiri::XML::Node.new(child_xpath, parent_node)
                      .tap { |created_child_node| parent_node << created_child_node }
                end

            created_node.set_attribute(attribute, encoded_value)
          end
            .first
        end
      end

      @attributes = []

      class << self
        attr_accessor :attributes

        def required_attributes
          attributes.select(&:required?)
        end

        def attribute(name, *selectors, required: false, converter: nil, validator: nil)
          new_attribute = Attribute.new(name, selectors, required: required, converter: converter, validator: validator)
          attributes << new_attribute

          define_method(name) do
            new_attribute.retrieve_from(self)
          end

          define_method("#{name}=") do |value|
            new_attribute.assign_to(self, value)
          end
        end

        def create(configuration, attributes = {})
          node = Nokogiri::XML::Node.new("w:style", configuration.styles_parent_node)
          configuration.styles_parent_node.add_child(node)

          Elements::Style.new(configuration, node, **attributes)
        end
      end

      def initialize(configuration, node, **attributes)
        @configuration = configuration
        @node = node

        attributes.each do |name, value|
          self.send("#{name}=", value)
        end
      end

      attr_accessor :node

      attribute :id, "./@w:styleId", required: true
      attribute :name, "./w:name/@w:val", "./w:next/@w:val", required: true
      attribute :type, ".//@w:type", required: true, validator: Validators::ValueValidator.new("paragraph", "character", "table", "numbering")
      attribute :keep_next, "./w:pPr/w:keepNext/@w:val", converter: Converters::BooleanConverter
      attribute :keep_lines, "./w:pPr/w:keepLines/@w:val", converter: Converters::BooleanConverter
      attribute :page_break_before, "./w:pPr/w:pageBreakBefore/@w:val", converter: Converters::BooleanConverter
      attribute :widow_control, "./w:pPr/w:widowControl/@w:val", converter: Converters::BooleanConverter
      attribute :shading_style, "./w:pPr/w:shd/@w:val", "./w:rPr/w:shd/@w:val"
      attribute :shading_color, "./w:pPr/w:shd/@w:color", "./w:rPr/w:shd/@w:color", validator: Validators::ColorValidator
      attribute :shading_fill, "./w:pPr/w:shd/@w:fill", "./w:rPr/w:shd/@w:fill"
      attribute :suppress_auto_hyphens, "./w:pPr/w:suppressAutoHyphens/@w:val", converter: Converters::BooleanConverter
      attribute :bidirectional_text, "./w:pPr/w:bidi/@w:val", converter: Converters::BooleanConverter
      attribute :spacing_before, "./w:pPr/w:spacing/@w:before"
      attribute :spacing_after, "./w:pPr/w:spacing/@w:after"
      attribute :line_spacing, "./w:pPr/w:spacing/@w:line"
      attribute :line_rule, "./w:pPr/w:spacing/@w:lineRule"
      attribute :indent_left, "./w:pPr/w:ind/@w:left"
      attribute :indent_right, "./w:pPr/w:ind/@w:right"
      attribute :indent_first_line, "./w:pPr/w:ind/@w:firstLine"
      attribute :align, "./w:pPr/w:jc/@w:val"
      attribute :font, "./w:rPr/w:rFonts/@w:ascii", "./w:rPr/w:rFonts/@w:cs", "./w:rPr/w:rFonts/@w:hAnsi", "./w:rPr/w:rFonts/@w:eastAsia" # setting :font, will set all other fonts
      attribute :font_ascii, "./w:rPr/w:rFonts/@w:ascii"
      attribute :font_cs, "./w:rPr/w:rFonts/@w:cs"
      attribute :font_hAnsi, "./w:rPr/w:rFonts/@w:hAnsi"
      attribute :font_eastAsia, "./w:rPr/w:rFonts/@w:eastAsia"
      attribute :bold, "./w:rPr/w:b/@w:val", "./w:rPr/w:bCs/@w:val", converter: Converters::BooleanConverter
      attribute :italic, "./w:rPr/w:i/@w:val", "./w:rPr/w:iCs/@w:val", converter: Converters::BooleanConverter
      attribute :caps, "./w:rPr/w:caps/@w:val", converter: Converters::BooleanConverter
      attribute :small_caps, "./w:rPr/w:smallCaps/@w:val", converter: Converters::BooleanConverter
      attribute :strike, "./w:rPr/w:strike/@w:val", converter: Converters::BooleanConverter
      attribute :double_strike, "./w:rPr/w:dstrike/@w:val", converter: Converters::BooleanConverter
      attribute :outline, "./w:rPr/w:outline/@w:val", converter: Converters::BooleanConverter
      attribute :outline_level, "./w:pPr/w:outlineLvl/@w:val"
      attribute :font_color, "./w:rPr/w:color/@w:val", validator: Validators::ColorValidator
      attribute :font_size, "./w:rPr/w:sz/@w:val", "./w:rPr/w:szCs/@w:val", converter: Converters::FontSizeConverter
      attribute :font_size_cs, "./w:rPr/w:szCs/@w:val", converter: Converters::FontSizeConverter
      attribute :underline_style, "./w:rPr/w:u/@w:val"
      attribute :underline_color, "./w:rPr/w:u/@w:color", validator: Validators::ColorValidator
      attribute :spacing, "./w:rPr/w:spacing/@w:val"
      attribute :kerning, "./w:rPr/w:kern/@w:val"
      attribute :position, "./w:rPr/w:position/@w:val"
      attribute :text_fill_color, "./w:rPr/w14:textFill/w14:solidFill/w14:srgbClr/@w14:val", validator: Validators::ColorValidator
      attribute :vertical_alignment, "./w:rPr/w:vertAlign/@w:val"
      attribute :lang, "./w:rPr/w:lang/@w:val"

      def valid?
        self.class.required_attributes.all? do |a|
          attribute_value = a.retrieve_from(self)

          a.validator&.validate(attribute_value)
        end
      end

      def to_xml
        node.to_xml
      end

      def remove
        node.remove
        @configuration.styles.delete(self)
      end
    end
  end
end
