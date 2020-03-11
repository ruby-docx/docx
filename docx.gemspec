$:.unshift File.expand_path("../lib", __FILE__)
require 'docx/version'

Gem::Specification.new do |s|
  s.name        = 'docx'
  s.version     = Docx::VERSION
  s.licenses    = ['MIT']
  s.summary     = 'a ruby library/gem for interacting with .docx files'
  s.description = 'thin wrapper around rubyzip and nokogiri as a way to get started with docx files'
  s.authors     = ['Christopher Hunt', 'Marcus Ortiz', 'Higgins Dragon', 'Toms Mikoss', 'Sebastian Wittenkamp']
  s.email       = ['chrahunt@gmail.com']
  s.homepage    = 'https://github.com/chrahunt/docx'
  s.files       = Dir["README.md", "LICENSE.md", "lib/**/*.rb"]
  s.required_ruby_version = '>= 2.4.0'

  s.add_dependency 'nokogiri', '~> 1.10', '>= 1.10.4'
  s.add_dependency 'rubyzip',  '~> 2.0'

  s.add_development_dependency 'rspec', '~> 3.7'
  s.add_development_dependency 'rake', '~> 13.0'
end
