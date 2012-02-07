require 'rubygems'
require 'bundler/setup'
require 'excelx_preview'
data = ExcelX::Previewer.preview(ARGV[0])["sheet1"]

