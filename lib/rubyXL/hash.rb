require 'rubygems'
require 'nokogiri'
# modified from http://stackoverflow.com/questions/1230741/convert-a-nokogiri-document-to-a-ruby-hash/1231297#1231297

module RubyXL
  class Hash < ::Hash

    def self.xml_node_to_hash_array(node)
      return {} if node.nil?
      self.xml_node_to_hash(node, :enclose_in_array)
    end

    # Slightly fixing this abomination of a code. Need to rethink the whole storage model of this stuff.
    def self.xml_node_to_hash(node, enclose_in_array = false)
      return prepare(node.content.to_s) unless node.element?

      # If we are at the root of the document, start the hash
      result_hash = {}

      unless node.attributes.empty?
        result_hash[:attributes] = {}

        node.attributes.keys.each { |key|
          result_hash[:attributes][node.attributes[key].name.to_sym] = prepare(node.attributes[key].value)
        }
      end

      node.children.each { |child|
        result = xml_node_to_hash(child)

        sym = child.name.to_sym

        if sym == :text then
          return prepare(result) unless child.next_sibling || child.previous_sibling
        elsif result_hash[sym]
          result_hash[sym] = [ result_hash[sym] ] unless result_hash[sym].is_a?(Object::Array)
          result_hash[sym] << prepare(result)
        else
          result_hash[sym] = enclose_in_array ? [ prepare(result) ] : prepare(result)
        end
      }

      result_hash
    end

    def self.prepare(data)
      (data.class == String && data.to_i.to_s == data) ? data.to_i : data
    end
  end
end
