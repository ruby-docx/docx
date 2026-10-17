# frozen_string_literal: true

require 'tmpdir'

module FixtureHelpers
  FIXTURES_DIR = File.expand_path('../fixtures', __dir__)

  def fixture_path(name)
    File.join(FIXTURES_DIR, name)
  end

  # A path inside a per-example temporary directory, so saved documents never
  # land in spec/fixtures and are always cleaned up.
  def tmp_path(name)
    File.join(@tmp_dir, name)
  end
end

RSpec.configure do |config|
  config.include FixtureHelpers

  config.around do |example|
    Dir.mktmpdir('docx_spec') do |dir|
      @tmp_dir = dir
      example.run
    end
  end
end
