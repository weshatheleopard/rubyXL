module RubyXL
  class GenericStorage < Hash

    def initialize(local_dir_path)
      @local_dir_path = local_dir_path
      super
    end

    def load(root_dir, mode = 'r')
      dirpath = File.join(root_dir, @local_dir_path)
      if File.directory?(dirpath) then
        (Dir.new(dirpath).entries - ['.', '..', '.DS_Store']).each { |filename|
          self[filename] = File.open(File.join(dirpath, filename), mode).read
        }
      end

      self
    end

    def add_to_zip(zipfile)
      each_pair { |filename, data|
        zipfile.get_output_stream(File.join(@local_dir_path, filename)) { |f| f.puts(data) }
      }
    end

  end
end
