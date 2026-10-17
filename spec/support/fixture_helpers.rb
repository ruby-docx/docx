# frozen_string_literal: true

require 'tmpdir'

module FixtureHelpers
  FIXTURES_DIR = File.expand_path('../fixtures', __dir__)

  def fixture_path(name)
    File.join(FIXTURES_DIR, name)
  end

  # The binary contents of fixture +name+ with the given zip entries left out,
  # for testing documents that lack a part (open it with Docx::Document.open).
  def fixture_without(name, *excluded_entries)
    Zip::File.open(fixture_path(name)) do |zip|
      Zip::OutputStream.write_buffer do |out|
        zip.each do |entry|
          next if !entry.file? || excluded_entries.include?(entry.name)

          out.put_next_entry(entry.name)
          out.write(zip.read(entry.name))
        end
      end.string
    end
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
