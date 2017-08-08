require './lib/batch_factory/version'

Gem::Specification.new do |s|
  s.files = Dir.glob("lib/**/*.rb")
  s.test_files  = Dir.glob("{spec,test}/**/*.rb")

  s.name        = "batch_factory"
  s.version     = BatchFactory::VERSION::STRING
  s.authors     = ["Denis Ivanov"]
  s.email       = ["visible@jumph4x.net"]

  s.summary     = "Meaningful spreadsheet access"
  s.description = "Assumes first row to be keys and returns hashes for consecutive rows"
  s.homepage    = "https://github.com/jumph4x/batch-factory"

  s.add_dependency 'roo', '~> 2.7.1'
  s.add_dependency 'roo-xls', '~> 1.1.0'
  s.add_dependency 'activesupport'

  s.add_development_dependency 'rspec', '~> 2.10'
  s.add_development_dependency 'rspec-its'
end

