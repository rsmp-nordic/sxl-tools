#!/usr/bin/env ruby
require 'rubyXL'
require 'rubyXL/convenience_methods'

# xlsx2csv reads sxl i excel format and outputs to csv-format

xlsx = ARGV[0].dup
workbook = RubyXL::Parser.parse(xlsx)

system("mkdir Objects")
workbook.each do |sheet|
  File.open("Objects/" + sheet.sheet_name + ".csv", "w") do |f|
    row = 0
    while sheet[row] != nil do
      line = ""
      col = 0
      while sheet[row][col] != nil do
        line = line + sheet[row][col].value.to_s + ";" if sheet[row][col].value != nil
        col = col + 1
      end
      f.write line
      row = row + 1
    end
  end
end

system("zip -q -j Objects.zip Objects/*")
system("rm -rf Objects")

# Rename after source file
xlsx.gsub!(/ /, '_')
xlsx.gsub!(/.xlsx/, '')
system("mv Objects.zip " + xlsx + ".zip")
