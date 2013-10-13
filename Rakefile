require 'rake'
require 'rspec/core/rake_task'
require 'bundler/gem_tasks'

RSpec::Core::RakeTask.new('spec')

desc 'Run tests.'
task default: :spec

desc "Open an irb session preloaded with this library."
task :console do
  sh "irb -I lib/ -r docx"
end
