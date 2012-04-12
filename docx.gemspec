$:.unshift File.expand_path("../lib", __FILE__)
require 'docx/version'

Gem::Specification.new do |s|
  s.name        = 'docx'
  s.version     = Docx::VERSION
  s.summary     = 'a ruby library/gem for interacting with .docx files'
  s.description = s.summary
  s.author      = 'Marcus Ortiz'
  s.email       = 'mportiz08@gmail.com'
  s.homepage    = 'https://github.com/mportiz08/docx'
  s.files       = Dir["README.md", "LICENSE.md", "lib/**/*.rb"]
  
  s.add_dependency 'nokogiri', '~> 1.5'
  s.add_dependency 'rubyzip',  '~> 0.9'
end
