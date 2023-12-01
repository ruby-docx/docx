# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Style do
  let(:fixture_path) { Dir.pwd + "/spec/fixtures/partial_styles/full.xml" }

  let(:node) do
    Nokogiri::XML(File.open(fixture_path)).root.children[1]
  end
  let(:style) { described_class.new(double(:configuration), node) }

  it "should extract attributes" do
    expect(style.id).to eq("Red")
  end

  describe "attribute getters" do
    it { expect(style.id).to eq("Red") }
    it { expect(style.name).to eq("Red") }
    it { expect(style.type).to eq("paragraph") }
    it { expect(style.keep_next).to eq(false) }
    it { expect(style.keep_lines).to eq(false) }
    it { expect(style.page_break_before).to eq(false) }
    it { expect(style.widow_control).to eq(true) }
    it { expect(style.suppress_auto_hyphens).to eq(false) }
    it { expect(style.bidirectional_text).to eq(false) }
    it { expect(style.spacing_before).to eq("0") }
    it { expect(style.spacing_after).to eq("200") }
    it { expect(style.line_spacing).to eq("240") }
    it { expect(style.line_rule).to eq("auto") }
    it { expect(style.indent_left).to eq("0") }
    it { expect(style.indent_right).to eq("0") }
    it { expect(style.indent_first_line).to eq("0") }
    it { expect(style.align).to eq("left") }
    it { expect(style.outline_level).to eq("9") }
    it { expect(style.font).to eq("Cambria") }
    it { expect(style.font_ascii).to eq("Cambria") }
    it { expect(style.font_cs).to eq("Arial Unicode MS") }
    it { expect(style.font_hAnsi).to eq("Cambria") }
    it { expect(style.font_eastAsia).to eq("Arial Unicode MS") }
    it { expect(style.bold).to eq(true) }
    it { expect(style.italic).to eq(false) }
    it { expect(style.caps).to eq(false) }
    it { expect(style.small_caps).to eq(false) }
    it { expect(style.strike).to eq(false) }
    it { expect(style.double_strike).to eq(false) }
    it { expect(style.outline).to eq(false) }
    it { expect(style.shading_style).to eq("clear") }
    it { expect(style.shading_color).to eq("auto") }
    it { expect(style.shading_fill).to eq("auto") } # TODO
    it { expect(style.font_color).to eq("99403d") }
    it { expect(style.font_size).to eq(12) }
    it { expect(style.font_size_cs).to eq(12) }
    it { expect(style.underline_style).to eq("none") }
    it { expect(style.underline_color).to eq("000000") }
    it { expect(style.spacing).to eq("0") }
    it { expect(style.kerning).to eq("0") }
    it { expect(style.position).to eq("0") }
    it { expect(style.text_fill_color).to eq("9A403E") }
    it { expect(style.vertical_alignment).to eq("baseline") }
    it { expect(style.lang).to eq("en-US") }
  end

  it "should allow setting simple attributes" do
    style.id = "Blue"

    # Get persisted to the style method
    expect(style.id).to eq("Blue")

    # Gets persisted to the ./node
    expect(node.at_xpath("./@w:styleId").value).to eq("Blue")
  end

  it "should allow setting complex attributes" do
    style.shading_style = "complex"

    # Get persisted to the style method
    expect(style.shading_style).to eq("complex")

    # Gets persisted to the node
    expect(node.at_xpath("./w:pPr/w:shd/@w:val").value).to eq("complex")
    expect(node.at_xpath("./w:rPr/w:shd/@w:val").value).to eq("complex")
  end

  describe "#to_xml" do
    it "should return the node as XML" do
      expect(style.to_xml).to eq(node.to_xml)
    end

    it "should change underlying XML when attributes are changed" do
      style.id = "blue"
      style.name = "Blue"

      expect(style.to_xml).to eq(node.to_xml)
      expect(style.to_xml).to include('<w:style w:type="paragraph" w:styleId="blue">')
      expect(style.to_xml).to include('<w:name w:val="Blue"/>')
      expect(style.to_xml).to include('<w:next w:val="Blue"/>')
    end
  end

  describe "basic" do
    let(:fixture_path) { Dir.pwd + "/spec/fixtures/partial_styles/basic.xml" }

    it "should allow setting simple attributes" do
      expect(style.id).to eq("MyCustomStyle")
      style.id = "Blue"

      # Get persisted to the style method
      expect(style.id).to eq("Blue")

      # Gets persisted to the node
      expect(node.at_xpath("./@w:styleId").value).to eq("Blue")
    end

    it "should allow setting complex attributes" do
      expect(style.shading_style).to eq(nil)
      expect(style.to_xml).to_not include('<w:shd w:val="complex"/>')
      style.shading_style = "complex"

      # Get persisted to the style method
      expect(style.shading_style).to eq("complex")

      # Gets persisted to the node
      expect(node.at_xpath("./w:pPr/w:shd/@w:val").value).to eq("complex")
      expect(node.at_xpath("./w:rPr/w:shd/@w:val").value).to eq("complex")
      expect(style.to_xml).to include('<w:shd w:val="complex"/>')
    end
  end
end