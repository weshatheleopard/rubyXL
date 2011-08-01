require 'rubygems'
require 'xmlsimple'
require 'zip/zip'
require 'FileUtils'
require 'zip/zipfilesystem'
require 'zip'

module RubyXL
#takes /Users/vbhagwat/Desktop/test1/Workbook1.xlsx
#copies, unzips, zips

# #takes path of directoy to be compressed
# #writes zip in that directory
# def compress(path)
#   puts 'path'
#   p path
#   path.sub!(%r[/$],'')
#   p path
#   archive = File.join(path,File.basename(path))+'.zip'
#   puts 'archive'
#   p archive
#   FileUtils.rm archive, :force=>true
#   Zip::ZipFile.open(archive, 'w') do |zipfile|
#     Dir["#{path}/**/**"].reject{|f|f==archive}.each do |file|
#       puts 'here'
#       p file.to_s()
#       temp = file
#       p temp.sub(path+'/','')
#       zipfile.add(file.sub(path+'/',''),file)
#     end
#   end
# end
# 
# #unzips file
# def unzip(zipPath,zipDirPath)
# Zip::ZipFile.open(zipPath) { |zip_file|
#   zip_file.each { |f|
#     fpath = File.join(zipDirPath, f.name)
#     FileUtils.mkdir_p(File.dirname(fpath))
#     zip_file.extract(f, fpath) unless File.exist?(fpath)
#   }
# }
# end

z = MyZip.new


dirPath = '/Users/vbhagwat/Desktop/test2/'
filePath = dirPath + 'Workbook1.xlsx'
zipDirPath = dirPath+'Workbook1/'
zipPath = zipDirPath + 'Workbook1.zip'

FileUtils.mkdir_p(dirPath+'Workbook1')
FileUtils.cp(filePath,zipPath)

z.unzip(zipPath,zipDirPath)
#TODO test if xml_in then xml_out corrupts it
# FileUtils.rm(zipPath) #removes zip file created

#commented out because i copy this anyway
# contents = XmlSimple.xml_in(zipDirPath + '[Content_Types].xml')
# contents = XmlSimple.xml_out(contents).gsub(/<\/?opt>/,'')
# file = File.new(zipDirPath+'[Content_Types].xml', 'w+')
# file.write(contents)
# file.close

# contents = XmlSimple.xml_in(zipDirPath+'docProps/app.xml')
# contents = XmlSimple.xml_out(contents).gsub(/<\/?opt>/,'')
# file = File.new(zipDirPath+'docProps/app.xml', 'w+')
# file.write(contents)
# file.close
# 
# contents = XmlSimple.xml_in(zipDirPath+'docProps/core.xml')
# contents = XmlSimple.xml_out(contents).gsub(/<\/?opt>/,'')
# file = File.new(zipDirPath+'docProps/core.xml', 'w+')
# file.write(contents)
# file.close
# 
# contents = XmlSimple.xml_in(zipDirPath+'xl/_rels/workbook.xml.rels')
# contents = XmlSimple.xml_out(contents).gsub(/<\/?opt>/,'')
# file = File.new(zipDirPath+'xl/_rels/workbook.xml.rels', 'w+')
# file.write(contents)
# file.close
# 
# contents = XmlSimple.xml_in(zipDirPath+'xl/styles.xml')
# contents = XmlSimple.xml_out(contents).gsub(/<\/?opt>/,'')
# file = File.new(zipDirPath+'xl/styles.xml', 'w+')
# file.write(contents)
# file.close
# 
# contents = XmlSimple.xml_in(zipDirPath+'xl/workbook.xml')
# contents = XmlSimple.xml_out(contents).gsub(/<\/?opt>/,'')
# file = File.new(zipDirPath+'xl/workbook.xml', 'w+')
# file.write(contents)
# file.close
# 
# #commented out because i copy this anyway
# # contents = XmlSimple.xml_in(zipDirPath+'xl/theme/theme1.xml')
# # contents = XmlSimple.xml_out(contents).gsub(/<\/?opt>/,'')
# # file = File.new(zipDirPath+'[Content_Types].xml', 'w+')
# # file.write(contents)
# # file.close
# 
# contents = XmlSimple.xml_in(zipDirPath+'xl/worksheets/sheet1.xml')
# contents = XmlSimple.xml_out(contents).gsub(/<\/?opt>/,'')
# file = File.new(zipDirPath+'xl/worksheets/sheet1.xml', 'w+')
# file.write(contents)
# file.close

#TODO manually reorder the xml tags to correlate, then compress using this compression method, 
#(not archive utility), then see if that xlsx file opens

z.compress('/Users/vbhagwat/Desktop/test2/Workbook1/')


# compress(dirPath)

end