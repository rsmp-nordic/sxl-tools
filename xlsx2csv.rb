#!/usr/bin/env ruby
require 'rubyXL'
require 'rubyXL/convenience_methods'

# xlsx2csv reads sxl i excel format and outputs to csv-format

xlsx = ARGV[0].dup
workbook = RubyXL::Parser.parse(xlsx)

system("mkdir Objects")
workbook.each do |sheet|
  File.open("Objects/" + sheet.sheet_name + ".csv", "w") do |f|

    # Get max rows and columns for each sheet
    max_column = 0
    row = 0
    rows = sheet.map {|row| row && row.cells.each { |cell| cell && cell.value != nil}}
    while row < rows.size
      last_column = rows.compact.max_by{|row| row.size}.size
      if last_column > max_column
        max_column = last_column
      end
      row = row + 1
    end
    max_column = max_column - 1

    row = 0
    while row < rows.size
      line = ""
      col = 0

      # Output the same number of columns
      while col < max_column do
        if sheet[row] != nil and sheet[row][col] != nil and sheet[row][col].value != nil
          line = line + sheet[row][col].value.to_s + ";"
        else
          line = line + ";"
        end
        col = col + 1
      end

      f.puts line 
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
