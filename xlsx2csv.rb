#!/usr/bin/env ruby
require 'rubyXL'
require 'rubyXL/convenience_methods'

# xlsx2csv reads sxl i excel format and outputs to csv-format

xlsx = ARGV[0].dup
workbook = RubyXL::Parser.parse(xlsx)

Dir.mkdir("Objects")
workbook.each do |sheet|
  oname = sheet.sheet_name.gsub(/ /, '_')

  File.open("Objects/" + oname + ".csv", "w") do |f|

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

    row = 0
    while row < rows.size
      col = 0

      # Output the same number of columns
      while col < max_column do
        f.print ';' unless col == 0
        if sheet[row] != nil and sheet[row][col] != nil and sheet[row][col].value != nil
          value = sheet[row][col].value.to_s
          # Excel adds an extra quotation mark if it finds one
          if value.match(/"/)
            value.gsub!(/"/, '""')
          end
          # Excel quotes the value if it contain newlines, semicolons or quotations marks
          if value.match(/\n/) or value.match(/;/) or value.match(/"/)
            value = '"' + value + '"'
          end

          # Older versions of Excel exports to CP-1252 instead of UTF-8
          value.encode!("CP1252")

          f.print value
        end
        col = col + 1
      end

      f.print "\r\n"
      row = row + 1
    end
  end
end
