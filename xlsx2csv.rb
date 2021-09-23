#!/usr/bin/env ruby
require 'rubyXL'
require 'rubyXL/convenience_methods'

# xlsx2csv reads sxl i excel format and outputs to csv-format

XLSX = ARGV[0]
workbook = RubyXL::Parser.parse(XLSX)

system("mkdir Objects")
workbook.each do |sheet|
  row = 0
  while sheet[row] != nil do
    line = ""
    col = 0
    while sheet[row][col] != nil do
      line = line + sheet[row][col].value.to_s + ";" if sheet[row][col].value != nil
      col = col + 1
    end
    puts line
    row = row + 1
  end
end

system("zip -q -j Objects.zip Objects/*")
system("rm -rf Objects")

# Rename after source file
XLSX.gsub!(/ /, '_')
XLSX.gsub!(/.xlsx/, '')
system("mv Objects.zip " + XLSX + ".zip")
