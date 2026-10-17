# frozen_string_literal: true

require 'spec_helper'
require 'docx'

describe Docx::SimpleInspect do
  let(:object) do
    Class.new do
      include Docx::SimpleInspect

      def initialize
        @short = 'abc'
        @long = 'x' * 500
      end
    end.new
  end

  describe '#inspect' do
    it 'lists every instance variable' do
      expect(object.inspect).to include('@short="abc"', '@long=')
    end

    it 'truncates long values' do
      expect(object.inspect).to include(%(@long="#{'x' * 100}...))
      expect(object.inspect.length).to be < 300
    end
  end
end
