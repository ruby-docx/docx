# frozen_string_literal: true

require 'spec_helper'
require 'docx'
require 'pathname'

# Docx::Document owns the .docx package: opening it, loading its parts
# (document, styles, relationships, headers, footers), exposing document-wide
# lookups, and writing it back out. Behavior of the individual elements is
# covered by the specs under spec/docx/containers and spec/docx/elements.
describe Docx::Document do
  describe '.open' do
    context 'with a String path' do
      let(:doc) { described_class.open(fixture_path('basic.docx')) }

      it_behaves_like 'reading basic.docx'
    end

    context 'with the binary contents of a file' do
      let(:doc) { described_class.open(File.binread(fixture_path('basic.docx'))) }

      it_behaves_like 'reading basic.docx'
    end

    # Regression test for #101: path-like objects work the same as a String path.
    context 'with a Pathname' do
      let(:doc) { described_class.open(Pathname.new(fixture_path('basic.docx'))) }

      it_behaves_like 'reading basic.docx'
    end

    it 'reads a main part named other than document.xml (e.g. Office365 document2.xml)' do
      doc = described_class.open(fixture_path('office365.docx'))

      expect(doc.paragraphs).not_to be_empty
    end

    it 'yields the document when given a block' do
      expect { |block| described_class.open(fixture_path('basic.docx'), &block) }
        .to yield_with_args(an_instance_of(described_class))
    end

    it 'raises when the file does not exist' do
      path = fixture_path('invalid_file_path.docx')

      expect { described_class.open(path) }.to raise_error(Zip::Error, "File #{path} not found")
    end
  end

  describe '#inspect' do
    it 'is short but mentions every instance variable' do
      doc = described_class.open(fixture_path('office365.docx'))

      expect(doc.inspect.length).to be < 1000
      doc.instance_variables.each do |var|
        expect(doc.inspect).to match(/#{var}/)
      end
    end
  end

  describe '#paragraphs' do
    it 'returns only the paragraphs directly in the body, not those inside tables' do
      doc = described_class.open(fixture_path('tables.docx'))

      expect(doc.paragraphs).to all(be_an_instance_of(Docx::Elements::Containers::Paragraph))
      expect(doc.paragraphs.map(&:text)).not_to include('Second table')
    end

    it 'gives paragraphs the document properties and the document itself' do
      doc = described_class.open(fixture_path('styles.docx'))

      expect(doc.paragraphs.first.font_size).to eq(doc.font_size)
      expect(doc.paragraphs.first.style).to eq('Title')
    end
  end

  describe '#each_paragraph' do
    it 'yields each paragraph' do
      doc = described_class.open(fixture_path('basic.docx'))

      expect { |block| doc.each_paragraph(&block) }.to yield_control.twice
    end
  end

  describe '#tables' do
    it 'returns the tables of the body' do
      doc = described_class.open(fixture_path('tables.docx'))

      expect(doc.tables.size).to eq(2)
      expect(doc.tables).to all(be_an_instance_of(Docx::Elements::Containers::Table))
    end
  end

  describe '#bookmarks' do
    it 'excludes the _GoBack bookmark Word adds automatically' do
      doc = described_class.open(fixture_path('basic.docx'))

      expect(doc.bookmarks.keys).not_to include('_GoBack')
    end

    it 'includes bookmarks found in headers and footers' do
      doc = described_class.open(fixture_path('multi_doc_bookmarks.docx'))

      expect(doc.bookmarks['header_bookmark']).to be_a(Docx::Elements::Bookmark)
      expect(doc.bookmarks['footer_bookmark']).to be_a(Docx::Elements::Bookmark)
    end
  end

  describe '#headers' do
    it 'maps each header part name to its XML' do
      doc = described_class.open(fixture_path('multi_doc.docx'))

      expect(doc.headers.keys).to eq(['header1'])
      expect(doc.headers['header1'].text).to eq('Hello from the header.')
    end
  end

  describe '#footers' do
    it 'maps each footer part name to its XML' do
      doc = described_class.open(fixture_path('multi_doc.docx'))

      expect(doc.footers.keys).to eq(['footer1'])
      expect(doc.footers['footer1'].text).to eq('Hello from the footer.')
    end
  end

  # ECMA-376 17.3.2.38: w:sz is in half-points.
  describe '#font_size' do
    it 'reads the default font size in points from styles.xml' do
      expect(described_class.open(fixture_path('formatting.docx')).font_size).to eq(11)
    end

    it 'is nil when the document has no styles.xml' do
      expect(described_class.open(fixture_path('no_styles.docx')).font_size).to be_nil
    end
  end

  describe '#hyperlinks' do
    it 'maps relationship ids to hyperlink targets' do
      doc = described_class.open(fixture_path('formatting.docx'))

      expect(doc.hyperlinks.values).to include('http://www.google.com/')
    end

    # Regression test for #158: relationships are loaded independently of
    # styles, so a document without word/styles.xml still has hyperlinks.
    it 'loads hyperlinks when styles.xml is missing' do
      doc = described_class.open(fixture_path('no_styles_with_hyperlink.docx'))

      expect(doc.hyperlinks).to eq('rId4' => 'http://www.google.com/')
    end

    # Regression test for #127
    it 'skips relationships missing an Id or Target' do
      doc = described_class.open(fixture_path('malformed_hyperlink_rels.docx'))

      expect { doc.hyperlinks }.not_to raise_error
      expect(doc.hyperlinks).to be_a(Hash)
    end

    # Regression tests for #192
    context 'when the document has no relationships part' do
      let(:doc) do
        described_class.open(fixture_without('basic.docx', 'word/_rels/document.xml.rels'))
      end

      it 'is empty' do
        expect(doc.hyperlinks).to eq({})
      end

      it 'still reads the document' do
        expect(doc.paragraphs.map(&:text)).to eq(%w[hello world])
        expect(doc.to_html).to include('hello')
      end
    end
  end

  describe '#document_properties' do
    it 'bundles the document-wide font size and hyperlinks' do
      doc = described_class.open(fixture_path('formatting.docx'))

      expect(doc.document_properties).to eq(font_size: doc.font_size, hyperlinks: doc.hyperlinks)
    end
  end

  describe '#default_paragraph_style' do
    it 'returns the name of the default paragraph style' do
      expect(described_class.open(fixture_path('styles.docx')).default_paragraph_style).to eq('Normal')
    end

    # Regression test for #127
    it 'is nil when there is no default paragraph style' do
      expect(described_class.open(fixture_path('no_default_paragraph_style.docx')).default_paragraph_style).to be_nil
    end
  end

  describe '#styles_configuration' do
    let(:doc) { described_class.open(fixture_path('styles.docx')) }

    it 'wraps styles.xml in a memoized StylesConfiguration' do
      expect(doc.styles_configuration).to be_a(Docx::Elements::Containers::StylesConfiguration)
      expect(doc.styles_configuration.size).to eq(37)
      expect(doc.styles_configuration).to equal(doc.styles_configuration)
    end

    it 'resolves style names through #style_name_of' do
      expect(doc.style_name_of('Heading1')).to eq('Heading 1')
    end
  end

  describe '#to_s' do
    it 'joins the paragraph texts with newlines' do
      expect(described_class.open(fixture_path('basic.docx')).to_s).to eq("hello\nworld")
    end

    it 'handles an unusually structured document' do
      expect(described_class.open(fixture_path('weird_docx.docx')).to_s).to be_a(String)
    end
  end

  describe '#to_html' do
    it 'joins the paragraph html with newlines' do
      doc = described_class.open(fixture_path('basic.docx'))

      expect(doc.to_html).to eq(doc.paragraphs.map(&:to_html).join("\n"))
    end

    it 'handles internal (anchor) links' do
      expect { described_class.open(fixture_path('internal-links.docx')).to_html }.not_to raise_error
    end
  end

  describe '#to_xml' do
    it 'returns the original main document XML' do
      doc = described_class.open(fixture_path('basic.docx'))
      doc.paragraphs.first.text = 'changed'

      expect(doc.to_xml).to be_a(Nokogiri::XML::Document)
      expect(doc.to_xml.text).to include('hello')
    end
  end

  describe '#save' do
    context 'from a file' do
      let(:doc) { described_class.open(fixture_path('saving.docx')) }

      it_behaves_like 'saving to file'
    end

    context 'from the binary contents of a file' do
      let(:doc) { described_class.open(File.binread(fixture_path('saving.docx'))) }

      it_behaves_like 'saving to file'
    end

    context 'from a WPS-modified file' do
      let(:doc) { described_class.open(fixture_path('saving_wps.docx')) }

      it_behaves_like 'saving to file'
    end

    it 'writes back edits to the body' do
      doc = described_class.open(fixture_path('basic.docx'))
      doc.paragraphs.first.text = 'edited'
      doc.save(tmp_path('saved.docx'))

      expect(described_class.open(tmp_path('saved.docx')).paragraphs.first.text).to eq('edited')
    end

    it 'writes back edits to headers and footers' do
      doc = described_class.open(fixture_path('multi_doc.docx'))
      doc.headers['header1'].at_xpath('//w:t').content = 'Edited header.'
      doc.footers['footer1'].at_xpath('//w:t').content = 'Edited footer.'
      doc.save(tmp_path('saved.docx'))

      reopened = described_class.open(tmp_path('saved.docx'))
      expect(reopened.headers['header1'].text).to eq('Edited header.')
      expect(reopened.footers['footer1'].text).to eq('Edited footer.')
    end

    it 'writes back edits to styles' do
      doc = described_class.open(fixture_path('styles.docx'))
      doc.styles_configuration.add_style('Red', font_color: 'FF0000')
      doc.save(tmp_path('saved.docx'))

      reopened = described_class.open(tmp_path('saved.docx'))
      expect(reopened.styles_configuration.style_of('Red').font_color).to eq('FF0000')
    end
  end

  describe '#stream' do
    let(:doc) { described_class.open(fixture_path('basic.docx')) }

    it 'returns a rewound StringIO to send over HTTP' do
      expect(doc.stream).to be_a(StringIO)
      expect(doc.stream.pos).to eq(0)
    end

    context 'when reopened' do
      let(:doc) { described_class.open(described_class.open(fixture_path('basic.docx')).stream) }

      it_behaves_like 'reading basic.docx'
    end
  end

  # Regression tests for #168: ZIP64 is disabled only while writing.
  describe 'ZIP64 global configuration' do
    let(:doc) { described_class.open(fixture_path('basic.docx')) }

    around do |example|
      previous = Zip.write_zip64_support
      Zip.write_zip64_support = true
      example.run
      Zip.write_zip64_support = previous
    end

    it 'does not leak the disabled setting after saving' do
      doc.save(tmp_path('saved.docx'))

      expect(Zip.write_zip64_support).to eq(true)
    end

    it 'does not leak the disabled setting after streaming' do
      doc.stream

      expect(Zip.write_zip64_support).to eq(true)
    end
  end

  describe '#replace_entry' do
    it 'replaces a part of the package when saving' do
      doc = described_class.open(fixture_path('replacement.docx'))
      image = File.binread(fixture_path('replacement.png'))

      doc.replace_entry('word/media/image1.png', image)
      doc.save(tmp_path('saved.docx'))

      expect(Zip::File.open(tmp_path('saved.docx')) { |zip| zip.read('word/media/image1.png') }).to eq(image)
    end
  end
end
