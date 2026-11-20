# -*- encoding: utf-8 -*-
# stub: rubyXL 3.4.40 ruby lib

Gem::Specification.new do |s|
  s.name = "rubyXL".freeze
  s.version = "3.4.40".freeze

  s.required_rubygems_version = Gem::Requirement.new(">= 0".freeze) if s.respond_to? :required_rubygems_version=
  s.require_paths = ["lib".freeze]
  s.authors = ["Vivek Bhagwat".freeze, "Wesha".freeze]
  s.date = "2026-11-20"
  s.description = "rubyXL is a gem which allows the parsing, creation, and manipulation of Microsoft Excel (.xlsx/.xlsm) Documents".freeze
  s.email = "bhagwat.vivek@gmail.com".freeze
  s.extra_rdoc_files = ["LICENSE.txt".freeze, "README.rdoc".freeze]
  s.files = ["LICENSE.txt".freeze, "README.rdoc".freeze, "Rakefile".freeze, "VERSION".freeze, "bin/console".freeze, "bin/setup".freeze, "lib/rubyXL.rb".freeze, "lib/rubyXL/cell.rb".freeze, "lib/rubyXL/convenience_methods.rb".freeze, "lib/rubyXL/convenience_methods/cell.rb".freeze, "lib/rubyXL/convenience_methods/color.rb".freeze, "lib/rubyXL/convenience_methods/font.rb".freeze, "lib/rubyXL/convenience_methods/workbook.rb".freeze, "lib/rubyXL/convenience_methods/worksheet.rb".freeze, "lib/rubyXL/objects/border.rb".freeze, "lib/rubyXL/objects/calculation_chain.rb".freeze, "lib/rubyXL/objects/cell_style.rb".freeze, "lib/rubyXL/objects/chartsheet.rb".freeze, "lib/rubyXL/objects/color.rb".freeze, "lib/rubyXL/objects/column_range.rb".freeze, "lib/rubyXL/objects/comments.rb".freeze, "lib/rubyXL/objects/connection.rb".freeze, "lib/rubyXL/objects/container_nodes.rb".freeze, "lib/rubyXL/objects/content_types.rb".freeze, "lib/rubyXL/objects/data_validation.rb".freeze, "lib/rubyXL/objects/document_properties.rb".freeze, "lib/rubyXL/objects/extensions.rb".freeze, "lib/rubyXL/objects/external_links.rb".freeze, "lib/rubyXL/objects/fill.rb".freeze, "lib/rubyXL/objects/filters.rb".freeze, "lib/rubyXL/objects/font.rb".freeze, "lib/rubyXL/objects/formula.rb".freeze, "lib/rubyXL/objects/ooxml_object.rb".freeze, "lib/rubyXL/objects/query_table.rb".freeze, "lib/rubyXL/objects/reference.rb".freeze, "lib/rubyXL/objects/relationships.rb".freeze, "lib/rubyXL/objects/root.rb".freeze, "lib/rubyXL/objects/shared_strings.rb".freeze, "lib/rubyXL/objects/sheet_common.rb".freeze, "lib/rubyXL/objects/sheet_data.rb".freeze, "lib/rubyXL/objects/simple_types.rb".freeze, "lib/rubyXL/objects/storage.rb".freeze, "lib/rubyXL/objects/stylesheet.rb".freeze, "lib/rubyXL/objects/text.rb".freeze, "lib/rubyXL/objects/theme.rb".freeze, "lib/rubyXL/objects/workbook.rb".freeze, "lib/rubyXL/objects/worksheet.rb".freeze, "lib/rubyXL/parser.rb".freeze, "lib/rubyXL/version.rb".freeze, "lib/rubyXL/worksheet.rb".freeze, "rubyXL.gemspec".freeze]
  s.homepage = "http://github.com/gilt/rubyXL".freeze
  s.licenses = ["MIT".freeze]
  s.rubygems_version = "3.6.9".freeze
  s.summary = "rubyXL is a gem which allows the parsing, creation, and manipulation of Microsoft Excel (.xlsx/.xlsm) Documents".freeze

  s.specification_version = 4

  s.add_runtime_dependency(%q<nokogiri>.freeze, [">= 1.10.8".freeze])
  s.add_runtime_dependency(%q<rubyzip>.freeze, [">= 3.2.2".freeze])
  s.add_development_dependency(%q<irb>.freeze, [">= 0".freeze])
  s.add_development_dependency(%q<rake>.freeze, [">= 0".freeze])
  s.add_development_dependency(%q<bundler>.freeze, [">= 0".freeze])
  s.add_development_dependency(%q<gem-release>.freeze, [">= 0".freeze])
  s.add_development_dependency(%q<bigdecimal>.freeze, [">= 0".freeze])
  s.add_development_dependency(%q<rspec>.freeze, [">= 0".freeze])
  s.add_development_dependency(%q<simplecov>.freeze, [">= 0".freeze])
  s.add_development_dependency(%q<stackprof>.freeze, [">= 0".freeze])
  s.add_development_dependency(%q<ruby-prof>.freeze, [">= 0".freeze])
  s.add_development_dependency(%q<rspec_junit_formatter>.freeze, [">= 0".freeze])
  s.add_development_dependency(%q<rubocop>.freeze, [">= 0".freeze])
end
