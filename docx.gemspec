$:.unshift File.expand_path("../lib", __FILE__)
require 'docx/version'

Gem::Specification.new do |s|
  s.name        = 'docx'
  s.version     = Docx::VERSION
  s.summary     = 'a ruby library/gem for interacting with .docx files'
  s.description = s.summary
  s.authors     = ['Christopher Hunt', 'Marcus Ortiz', 'Higgins Dragon', 'Toms Mikoss', 'Sebastian Wittenkamp']
  s.email       = ['sebastian@bitops.io']
  s.homepage    = 'https://github.com/bitops/docx'
  s.files       = Dir["README.md", "LICENSE.md", "lib/**/*.rb"]

  s.add_dependency 'nokogiri', '~> 1.5'
  s.add_dependency 'rubyzip',  '~> 1.1.6'

  s.add_development_dependency 'rspec'
end
