#!ruby -Ks
# encoding: utf-8
require 'win32ole'
require 'fileutils'

# This class converts visio file into pdf file.
class Visio2Pdf
  VSDEXTS = '.vsd'.freeze
  PDFEXTS = '.pdf'.freeze
  @visio = nil
  @in_dir = nil
  @vsd_fullpath = nil
  @pdf_fullpath = nil

  def visio2pdf_controller
    if ARGV.length == 2
      @in_dir, @_out_dir = ARGV
    else
      return
    end
    visio2pdf
  end

  def visio2pdf
    @visio = WIN32OLE.new('Visio.Application')
    @visio.alertresponse = 7
    files = Dir["#{@in_dir}/[^~]*{#{VSDEXTS}}"]
    files.sort.each do |file|
      get_filepath file
      visio2pdf_exec
    end
  ensure
    @visio.Quit
  end

  def visio2pdf_exec
    vsd = @visio.Documents.Open(@vsd_fullpath)
    vsd.ExportAsFixedFormat(
      FixedFormat: 1, OutputFileName: @pdf_fullpath, Intent: 0, PrintRange: 0
    )
  rescue
    puts "fail: #{@vsd_fullpath}"
  ensure
    vsd.Close
  end

  def get_filepath(file)
    @vsd_fullpath = File.expand_path(file)
    @pdf_fullpath = @vsd_fullpath.gsub(/\.vsd$/, PDFEXTS)
  end
end

exec = Visio2Pdf.new
exec.visio2pdf_controller
