#!ruby -Ks
# -*- coding: utf-8 -*-
require 'win32ole'
require 'fileutils'

def visio2pdf
  if ARGV.length == 2
    in_dir, out_dir = ARGV
  else
    return
  end
  visio = WIN32OLE.new('Visio.Application')
  exts = ['.vsd']
  begin
    files = Dir["#{in_dir}/[^~]*{#{exts}}"]
    files.sort.each_with_index do |file, i|
      vsd_fullpath = File.expand_path(file)
      pdf_fullpath = vsd_fullpath.gsub(/\.vsd$/, '.pdf')
      begin
        vsd = visio.Documents.Open(vsd_fullpath)
        vsd.ExportAsFixedFormat(
          'FixedFormat' => 1,
          'OutputFileName' => pdf_fullpath,
          'Intent' => 0,
          'PrintRange' => 0
        )
        puts "success: #{vsd_fullpath}"
      rescue
        puts "fail: #{vsd_fullpath}"
      ensure
        vsd.Close
      end
    end
  ensure
    visio.Quit
  end
end

visio2pdf
