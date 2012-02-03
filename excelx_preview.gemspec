# -*- encoding: utf-8 -*-
$:.push File.expand_path("../lib", __FILE__)
require "excelx_preview/version"

Gem::Specification.new do |s|
  s.name = "excelx_preview"
  s.version = ExcelxPreview::VERSION
  s.platform = Gem::Platform::RUBY
  s.authors = ["Krishnaprasad T Nair"]
  s.email = ["krishnaprasad@mobme.in"]
  s.homepage = ""
  s.summary = %q{Gem that extracts first 10 rows from an excelx file}
  s.description = %q{Gem that extracts first 10 rows from an excelx file.Supports simple formulas.}

  s.add_dependency  "nokogiri"
  s.add_dependency  'activesupport'
  s.add_dependency  'uuid'
  s.files = `git ls-files`.split("\n")
  # s.test_files    = `git ls-files -- {test,spec,features}/*`.split("\n")
  # s.executables   = `git ls-files -- bin/*`.split("\n").map{ |f| File.basename(f) }
  s.require_paths = ["lib"]
end