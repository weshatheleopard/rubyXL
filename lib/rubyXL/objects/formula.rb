require 'rubyXL/objects/ooxml_object'
require 'rubyXL/objects/simple_types'

module RubyXL

  NUMBER_REGEXP = /^-?\d+(\.\d+(?:e[+-]\d+)?)?/i
#  FLOAT_REGEXP  = /^\d+(?:\.\d+)?/

  # http://www.datypic.com/sc/ooxml/e-ssml_f-1.html

  class Formula < OOXMLObject
    define_attribute(:_,    :string, :accessor => :expression)
    define_attribute(:t,    RubyXL::ST_CellFormulaType, :default => 'normal')
    define_attribute(:aca,  :bool,   :default => false)
    define_attribute(:ref,  :ref)
    define_attribute(:dt2D, :bool,   :default => false)
    define_attribute(:dtr,  :bool,   :default => false)
    define_attribute(:del1, :bool,   :default => false)
    define_attribute(:del2, :bool,   :default => false)
    define_attribute(:r1,   :ref)
    define_attribute(:r2,   :ref)
    define_attribute(:ca,   :bool,   :default => false)
    define_attribute(:si,   :int)
    define_attribute(:bx,   :bool,   :default => false)
    define_element_name 'f'

    def calculate(calc_id)

    end

    def self.parse_formula(str)
      res = []

      str = str.lstrip

      until str == ''
        if str =~ /^\(/ then
puts "Expression detected: #{$'}"
          res2, str = parse_formula($')
          res << res2
          puts "--> result: #{res2}   remainder = [#{str}]"
        elsif str =~ /^\)/ then
puts "End of expression detected, res=#{res} remainder=#{$'}"
          return res, $'
        else
          token, str = parse_token(str)
          puts "--> token: #{token}   remainder = [#{str}]"
          res << token
        end
      end

      res
    end



    def self.parse_token(str)
      str = str.lstrip
      # Output is in the form: [ <operation>, <operand1>[, <operand2>[, <operand3>....]]
      res = []



      if str =~ NUMBER_REGEXP then # Number
puts "number detected, remainder = #{$'}"

        return $&.to_f, $'
      else
        if str=~ /[a-z]+[0-9]+/i then # possibly Reference
          remainder = $' # Must save first, because the next line will destroy the values.
          ref = RubyXL::Reference.new($&)
          return ref, remainder if (ref.last_row <= 1048576) && (ref.last_col <= 16383)
        end

        if str =~ /([a-z]+)\(/i then # possibly function
          # TODO: implement the list
          res[0] = $1

          loop do
            str = lstrip(str)
            r, str = parse($`)
            res << r

            str = lstrip(str)
            if str =~ /^\s*[,)]/ then
              str = $'
              break if $1 == ')'
            end
          end

        end

        if str =~ /[a-z_\\]([a-z0-9._])*/i then # possibly DefinedName
          # TODO:
          # * cannot be == "C", "c", "R", or "r", which are shorthands for selecting a row or column for the currently selected cell
          # * up to 255 chars
          dn = defined_names.find_by_name($&)
          return dn, $' unless dn.nil?        
        end

        # * operator: + – * / % ^ = > < >= <= <> &
        if str =~ />=|<=|<>|[-+*\/%^=><&]/ then
          return $&, $'
        end
   
      end

      # * function: "xxxx(...)"
      # * name:    
      # ** cannot be the same as cell refs

    end
  end

end
