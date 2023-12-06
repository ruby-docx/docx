# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Style do
  let(:fixture_path) { Dir.pwd + "/spec/fixtures/partial_styles/full.xml" }
  let(:fixture_xml) { File.read(fixture_path) }
  let(:node) { Nokogiri::XML(fixture_xml).root.children[1] }
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

  it "should allow setting attributes to nil" do
    style.shading_style = nil

    expect(style.shading_style).to eq(nil)
    expect(node.at_xpath("./w:pPr/w:shd/@w:val")).to eq(nil)
    expect { node.at_xpath("./w:pPr/w:shd/@w:val").value }.to raise_error(NoMethodError) # i.e. it's gone!
  end

  describe "#to_xml" do
    it "should return the node as XML" do
      expect(style.to_xml).to eq(node.to_xml)
    end

    it "should change underlying XML when attributes are changed" do
      style.id = "blue"
      style.name = "Blue"
      style.font_size = 20
      style.font_color = "0000FF"

      expect(style.to_xml).to eq(node.to_xml)
      expect(style.to_xml).to include('<w:style w:type="paragraph" w:styleId="blue">')
      expect(style.to_xml).to include('<w:name w:val="Blue"/>')
      expect(style.to_xml).to include('<w:next w:val="Blue"/>')
      expect(style.to_xml).to include('<w:sz w:val="40"/>')
      expect(style.to_xml).to include('<w:szCs w:val="40"/>')
      expect(style.to_xml).to include('<w:color w:val="0000FF"/>')
    end
  end

  describe "validation" do
    let(:fixture_path) { Dir.pwd + "/spec/fixtures/partial_styles/basic.xml" }

    it "validation: id" do
      expect { style.id = nil }.to raise_error(Docx::Errors::StyleRequiredPropertyValue)
    end

    it "validation: name" do
      expect { style.name = nil }.to raise_error(Docx::Errors::StyleRequiredPropertyValue)
    end

    it "validation: type" do
      expect { style.type = nil }.to raise_error(Docx::Errors::StyleRequiredPropertyValue)

      expect { style.type = "invalid" }.to raise_error(Docx::Errors::StyleInvalidPropertyValue)
    end

    it "true" do
      expect(style).to be_valid
    end

    describe "unhappy" do
      let(:fixture_xml) do
        <<~XML
        <?xml version="1.0" encoding="UTF-8"?>
        <w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
          <w:style w:type="" w:styleId="">
            <w:name/>
          </w:style>
        </w:styles>
        XML
      end

      it "false" do
        expect(style).to_not be_valid
      end
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