# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::Elements::Bookmark do
  def bookmark_in(body)
    described_class.new(body.at_xpath('.//w:bookmarkStart'))
  end

  def paragraph_texts(body)
    body.xpath('w:p').map(&:text)
  end

  it 'is backed by the w:bookmarkStart tag' do
    expect(described_class.tag).to eq('bookmarkStart')
  end

  it 'reads its name' do
    body = w_body('<w:p><w:bookmarkStart w:id="0" w:name="my_bookmark"/></w:p>')

    expect(bookmark_in(body).name).to eq('my_bookmark')
  end

  context 'inside a paragraph' do
    describe '#insert_text_before' do
      it 'appends to the run just before the bookmark' do
        body = w_body(<<~XML)
          <w:p><w:r><w:t>a</w:t></w:r><w:r><w:t>test text</w:t></w:r><w:bookmarkStart w:id="0" w:name="b"/><w:r><w:t>!</w:t></w:r></w:p>
        XML

        bookmark_in(body).insert_text_before('bar')

        expect(paragraph_texts(body)).to eq(['atest textbar!'])
      end

      it 'creates a run before the bookmark when none precedes it' do
        body = w_body('<w:p><w:bookmarkStart w:id="0" w:name="b"/><w:r><w:t>test text</w:t></w:r></w:p>')

        bookmark_in(body).insert_text_before('foo')

        expect(paragraph_texts(body)).to eq(['footest text'])
      end
    end

    describe '#insert_text_after' do
      it 'prepends to the run just after the bookmark' do
        body = w_body('<w:p><w:bookmarkStart w:id="0" w:name="b"/><w:r><w:t>test text</w:t></w:r><w:r><w:t>!</w:t></w:r></w:p>')

        bookmark_in(body).insert_text_after('foo')

        expect(paragraph_texts(body)).to eq(['footest text!'])
      end

      it 'creates a run after the bookmark when none follows it' do
        body = w_body('<w:p><w:r><w:t>test text</w:t></w:r><w:bookmarkStart w:id="0" w:name="b"/></w:p>')

        bookmark_in(body).insert_text_after('bar')

        expect(paragraph_texts(body)).to eq(['test textbar'])
      end
    end

    describe '#insert_multiple_lines' do
      it 'replaces the paragraph text and adds a paragraph per extra line' do
        body = w_body(<<~XML)
          <w:p><w:r><w:t>before</w:t></w:r></w:p>
          <w:p><w:r><w:t>place</w:t></w:r><w:bookmarkStart w:id="0" w:name="b"/><w:r><w:t>holder text</w:t></w:r></w:p>
          <w:p><w:r><w:t>after</w:t></w:r></w:p>
        XML

        bookmark_in(body).insert_multiple_lines(['line 1', 'line 2', 'line 3'])

        expect(paragraph_texts(body)).to eq(['before', 'line 1', 'line 2', 'line 3', 'after'])
      end
    end
  end

  # Regression tests for #111: some generators (e.g. Google Docs) place
  # w:bookmarkStart/End directly under w:body, outside any w:p. Inserted text
  # must land inside a paragraph, otherwise the run is orphaned and lost.
  context 'outside a paragraph (block-level)' do
    describe '#insert_text_after' do
      it 'inserts at the start of the following paragraph, after its properties' do
        body = w_body(<<~XML)
          <w:bookmarkStart w:id="0" w:name="b"/><w:bookmarkEnd w:id="0"/>
          <w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>bookmark_1</w:t></w:r></w:p>
        XML

        bookmark_in(body).insert_text_after('INSERTED ')

        expect(paragraph_texts(body)).to eq(['INSERTED bookmark_1'])
        expect(body.at_xpath('w:p').element_children.first.name).to eq('pPr')
      end

      it 'inserts at the start of a following paragraph without properties' do
        body = w_body('<w:bookmarkStart w:id="0" w:name="b"/><w:p><w:r><w:t>bookmark_1</w:t></w:r></w:p>')

        bookmark_in(body).insert_text_after('INSERTED ')

        expect(paragraph_texts(body)).to eq(['INSERTED bookmark_1'])
      end

      it 'creates a new paragraph after the bookmark when no paragraph follows' do
        body = w_body('<w:p><w:r><w:t>before</w:t></w:r></w:p><w:bookmarkStart w:id="0" w:name="b"/>')

        bookmark_in(body).insert_text_after('INSERTED')

        expect(paragraph_texts(body)).to eq(%w[before INSERTED])
      end
    end

    describe '#insert_text_before' do
      it 'appends to the end of the preceding paragraph' do
        body = w_body('<w:p><w:r><w:t>before</w:t></w:r></w:p><w:bookmarkStart w:id="0" w:name="b"/><w:p><w:r><w:t>after</w:t></w:r></w:p>')

        bookmark_in(body).insert_text_before(' BEFORE')

        expect(paragraph_texts(body)).to eq(['before BEFORE', 'after'])
      end

      it 'creates a new leading paragraph when no paragraph precedes' do
        body = w_body('<w:bookmarkStart w:id="0" w:name="b"/><w:p><w:r><w:t>bookmark_1</w:t></w:r></w:p>')

        bookmark_in(body).insert_text_before('BEFORE')

        expect(paragraph_texts(body)).to eq(%w[BEFORE bookmark_1])
      end
    end

    describe '#insert_multiple_lines' do
      it 'inserts new paragraphs at the bookmark without touching the following paragraph' do
        body = w_body('<w:bookmarkStart w:id="0" w:name="b"/><w:p><w:r><w:t>bookmark_1</w:t></w:r></w:p>')

        bookmark_in(body).insert_multiple_lines(%w[line1 line2 line3])

        expect(paragraph_texts(body)).to eq(%w[line1 line2 line3 bookmark_1])
      end
    end
  end
end
