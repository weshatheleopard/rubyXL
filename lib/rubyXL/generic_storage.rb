module RubyXL
  class GenericStorage < Hash
    attr_reader :local_dir_path

    def initialize(local_dir_path)
      @local_dir_path = local_dir_path
      @mode = 'r'
      super
    end

    def binary
      @mode = 'rb'
      self
    end

    def load_dir(root_dir)
      dirpath = File.join(root_dir, @local_dir_path)
      if File.directory?(dirpath) then
        (Dir.new(dirpath).entries - ['.', '..', '.DS_Store', '_rels']).each { |filename|
          # Making sure that the file will be automatically closed immediately after it has been read
          self[filename] = File.open(File.join(dirpath, filename), @mode) { |f| f.read }
        }
      end

      self
    end

    def load_file(root_dir, filename)
      filepath = File.join(root_dir, @local_dir_path, filename)
      # Making sure that the file will be automatically closed immediately after it has been read
      self[filename] = (File.open(filepath, @mode) { |f| f.read }) if File.readable?(filepath)
      self
    end

    def add_to_zip(zipfile)
      each_pair { |filename, data|
        zipfile.get_output_stream(File.join(@local_dir_path, filename)) { |f| f << data }
      }
    end

  end
end
