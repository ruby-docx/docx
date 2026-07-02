# frozen_string_literal: true

require 'spec_helper'
require 'docx/document'

describe 'Paragraph placeholder consolidation' do
  let(:described_class) { Docx::Elements::Containers::Paragraph }

  # Builds a bare <w:p> node so placeholder consolidation can be exercised
  # without a full .docx fixture.
  def paragraph_from(inner_runs)
    xml = <<~XML
      <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        #{inner_runs}
      </w:p>
    XML
    node = Nokogiri::XML(xml).root
    described_class.new(node)
  end

  describe '#validate_placeholder_content' do
    # Reproduces the real customer template: Word splits each {{token}} across
    # runs (proofing marks) and fuses one token's closing "}}" into the same run
    # as the next token's opening "{{" (e.g. a run reading "}} at CTC {{"). Two
    # placeholders then share that boundary run.
    let(:shared_boundary_runs) do
      <<~XML
        <w:r><w:t xml:space="preserve">Your title is {{</w:t></w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r><w:t>offer.job_title</w:t></w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r><w:t xml:space="preserve">}} at CTC {{</w:t></w:r>
        <w:proofErr w:type="spellStart"/>
        <w:r><w:t>offer.salary_amount</w:t></w:r>
        <w:proofErr w:type="spellEnd"/>
        <w:r><w:t>}}.</w:t></w:r>
      XML
    end

    it 'consolidates without altering the visible text' do
      paragraph = paragraph_from(shared_boundary_runs)

      expect(paragraph.text).to eq('Your title is {{offer.job_title}} at CTC {{offer.salary_amount}}.')
    end

    it 'keeps each placeholder within a single run so substitution leaves no orphan brace' do
      paragraph = paragraph_from(shared_boundary_runs)

      paragraph.each_text_run do |run|
        run.substitute('{{offer.job_title}}', 'Senior Account Executive')
        run.substitute('{{offer.salary_amount}}', '44,00,000')
      end

      expect(paragraph.text).to eq('Your title is Senior Account Executive at CTC 44,00,000.')
    end
  end
end
