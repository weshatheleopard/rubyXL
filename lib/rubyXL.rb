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

require 'zip'

# I do not appreciate the following hackery, but you have the developers of +zip+ and +rubyzip+
# gems to thank for it, as they *both* have choosen to call their projects' base classes +Zip+,
# and as a result, if +rubyXL+ is used in the project that also utilizes +zip+, it fails to work
# properly as it picks up the wrong class. So I have no choice but to do the cleanup of their mess
# in my code.
unless ::Object.const_defined?(:RubyZip, false)
  if ::Object.const_defined?(:Zip, false) then
    if ::Zip.const_defined?(:File, false) then
      # puts "DEBUG: RubyZip detected"
      ::RubyZip = ::Zip
    else
      # puts "DEBUG: Conflicting Zip detected"
      zip_backup = ::Zip
      ::Zip = nil
      require 'rubygems'
      gem 'rubyzip', :require => 'zip'
      ::RubyZip = ::Zip
      ::Zip = zip_backup
    end
  else
    # puts "DEBUG: Zip not detected at all"
    require 'rubygems'
    gem 'rubyzip', :require => 'zip'
    ::RubyZip = ::Zip
  end
end

module RubyXL
end
