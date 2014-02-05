require 'rubyXL/workbook'
require 'rubyXL/worksheet'
require 'rubyXL/cell'
require 'rubyXL/objects/reference'
require 'rubyXL/objects/column_range'
require 'rubyXL/objects/stylesheet'
require 'rubyXL/objects/shared_strings'
require 'rubyXL/objects/worksheet'
require 'rubyXL/objects/calculation_chain'
require 'rubyXL/objects/workbook'
require 'rubyXL/objects/document_properties'
require 'rubyXL/objects/relationships'
require 'rubyXL/parser'

if ::Object.const_defined?(:Zip) then
  if ::Zip.const_defined?(:File) then
    #puts "DEBUG: RubyZip detected"
    ::RubyZip = ::Zip
  else
    #puts "DEBUG: Conflicting Zip detected"
    zip_backup = ::Zip
    ::Zip = nil
    require 'rubygems'
    gem 'rubyzip'
    require 'zip'
    ::RubyZip = ::Zip
    ::Zip = zip_backup
  end
else
  #puts "DEBUG: No Zip detected"
  require 'rubygems'
  gem 'rubyzip'
  require 'zip'
  ::RubyZip = ::Zip
end

module RubyXL
end
