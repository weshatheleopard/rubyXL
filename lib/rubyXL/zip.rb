require 'rubygems'
require 'zip/zip'
require 'zip/zipfilesystem'

module RubyXL
  class MyZip

    # Unzips .zip file at zipPath to zipDirPath
    def unzip(zipPath,zipDirPath)
      Zip::ZipFile.open(zipPath) { |zip_file|
        zip_file.each { |f|
          fpath = File.join(zipDirPath, f.name)
          FileUtils.mkdir_p(File.dirname(fpath))
          zip_file.extract(f, fpath) unless File.exist?(fpath)
        }
      }
    end

  end
end
