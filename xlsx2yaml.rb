#!/usr/bin/env ruby
require 'yaml'
require 'rubyXL'
require 'rubyXL/convenience_methods'
require 'optparse'

# xlsx2yaml reads sxl in excel-format and outputs to yaml-format

# Get an individual object from the object sheet
def get_object(sheet, y)
  return if sheet[y] == nil
  return if sheet[y][0] == nil
  return if sheet[y][0].value == nil

  # Empty description/object field
  if sheet[y][1] == nil
    return [ sheet[y][0].value, nil, nil, nil, nil ]
  end

  # Empty component id
  if sheet[y][2] == nil
    return [ sheet[y][0].value, sheet[y][1].value, nil, nil, nil ]
  end

  # Empty NTSObjectId
  if sheet[y][3] == nil
    return [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value, nil, nil ]
  end

  # Empty externalNtsId
  if sheet[y][4] == nil
    return [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value, sheet[y][3].value, nil ]
  end

  [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value, sheet[y][3].value, sheet[y][4].value ]
end

# Get objects from a given section in the object sheet
def get_object_section(sheet, y, options)
  objects = {}
  while(object = get_object(sheet, y)) do
    key = object[0]  # e.g. Traffic Controller, Signal Group

    # Check if object type already exists, merge otherwise
    unless objects[key]
      # Object types has no componentId
      if object[2] == nil
        objects[key] = { 'description' => object[1] }
      else
        if options[:extended]
          objects[key] = { object[1] => { 'componentId' => object[2], 'ntsObjectId' => object[3] } }
          objects[key][object[1]].store("externalNtsId", to_integer(object[4])) if object[4] != nil
        else
          objects[key] = { object[1] => object[2] }
        end
      end
    else
        if options[:extended]
          newobject = { object[1] => { 'componentId' => object[2], 'ntsObjectId' => object[3] } }
          newobject[object[1]].store("externalNtsId", to_integer(object[4])) if object[4] != nil
        else
          newobject = { object[1] => object[2] }
        end
      objects[key] = objects[key].merge(newobject)
    end
    y = y + 1
  end
  objects
end

# Find command sections
def get_command_section(sheet)
  y = 4 # Section won't start before row 4
  return if sheet[y] == nil
  return if sheet[y][0] == nil
  return if sheet[y][0].value == nil

  sections = []
  while(y<100) do
    if sheet[y] != nil and sheet[y][0] != nil
      sections.append(y+2) if sheet[y][0].value.eql?("Functional position")
      sections.append(y+2) if sheet[y][0].value.eql?("Functional state")
      sections.append(y+2) if sheet[y][0].value.eql?("Manouver")
      sections.append(y+2) if sheet[y][0].value.eql?("Parameter")
    end
    y = y + 1
  end
  sections
end

# Get value from a description field
# E.g. 1: value1
#      2: value2
def get_value(field, key)
  value_pairs = field.split("\n")
  value_pairs.each {|v|
    k, val = v.split(": ")

    next if val.nil?
    if key == k
      # Remove pair from field
      replace = k + ": " + val
      field.gsub! replace, ''
      while field.chomp! do
        field.chomp!
      end
      return val
    end
  }
  return nil
end

def to_integer(val)
  v = Integer(val) rescue false
  if v == false
    v = val
  end
  v
end

# Re-indent code due to ruby yaml output fails to indent multiline text
# properly when blank lines occurs in the middle of text
def reindent(text)
  prev_line = ""
  prev_indent = ""
  text.each_line do |line|
    if line == "\n"
      # Check white spaces in the beginning of previous line
      indent =  prev_line.match(/^([ ]+)/)[0] rescue false
      if indent == false
        indent = prev_indent
      end
      line = indent + "\n"
    end
    print line
    prev_line = line
  end
end

options = {}
usage = "Usage: xlsx2yaml.rb [options] [XLSX]"
OptionParser.new do |opts|
  opts.banner = usage

  opts.on("-s", "--site", "Include site") do |s|
    options[:site] = s
  end

  opts.on("-e", "--extended", "Extended fields") do |e|
    options[:extended] = e
  end
end.parse!

if ARGV.length < 1
  puts usage
  exit 1
end

XLSX = ARGV[0]
sxl = Hash.new
workbook = RubyXL::Parser.parse(XLSX)

sites = {}
workbook.each do |sheet|
  case sheet.sheet_name
  when "Version"
    if options[:site]
      sxl["id"] = sheet[3][1].value if sheet[3]
      sxl["version"] = sheet[20][1].value
      sxl["date"] = sheet[20][2].value
      sxl["description"] = sheet[5][1].value
      if options[:extended]
        sxl["constructor"] = sheet[9][1].value if sheet[9]
        sxl["reviewed"] = sheet[11][1].value if sheet[11][1].value
        sxl["approved"] = sheet[13][1].value if sheet[13][1].value
        sxl["created-date"] = sheet[17][1].value if sheet[17]
        sxl["rsmp-version"] = sheet[25][1].value if sheet[25]
      end
    end
  when "Object types"
    # grouped objects
    sxl["objects"] = get_object_section(sheet, 6, options)

    # single objects
    sxl["objects"] = sxl["objects"].merge(get_object_section(sheet, 18, options))
  when "Alarms"
    y = 6
    while(sheet[y][0] != nil and sheet[y][0].value != nil) do
      # Get basic alarm info
      a = [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value,
            sheet[y][3].value, sheet[y][4].value, sheet[y][5].value,
            sheet[y][6].value, sheet[y][7].value]

      # Get each argument
      x = 8
      rv = {}
      while(sheet[y][x] != nil and sheet[y][x].value != nil) do
        rv[sheet[y][x].value] = {
            'type' => sheet[y][x+1].value,
            'description' => sheet[y][x+3].value
        }

        # No need to output values if type is boolean
        unless rv[sheet[y][x].value]['type'] == 'boolean'
          # Output values in a different way
          if sheet[y][x+2].value.start_with?("-")
            values = sheet[y][x+2].value.split("-")
            values.shift
            values.each {|v|
              v.delete!("\n")

              # Try to find the corresponding description in
              # the description field using "key: value" format
              # Removes the key in description on match
              field = sheet[y][x+3].value
              desc = get_value(field, v)
              if desc.nil? then
                desc = ''
              end

              # Add to yaml
              v = to_integer(v)
              if rv[sheet[y][x].value]['values']
                rv[sheet[y][x].value]['values'][v] = desc
              else
                rv[sheet[y][x].value]['values'] = {v => desc}
              end
            }
          else
            if options[:extended]
              rv[sheet[y][x].value]['range'] = sheet[y][x+2].value
            end
          end
        end

        # Remove the description field if it is empty
        if rv[sheet[y][x].value]['description'].empty?
          rv[sheet[y][x].value].delete('description')
        end

        x = x + 4
      end

      alarm = {
        'description' => a[3],
        'priority' => a[6],
        'category' => a[7]
      }
      alarm.store("object", a[1]) if a[1] != nil
      alarm.store("externalAlarmCodeId", a[4]) if a[4] != nil
      alarm.store("externalNtsAlarmCodeId", to_integer(a[5])) if a[5] != nil

      if !rv.empty?
        alarm['arguments'] = rv
      end

      # Add to yaml
      if sxl["objects"][a[0]]
        if sxl["objects"][a[0]]["alarms"]
          sxl["objects"][a[0]]["alarms"][a[2]] = alarm
        else
          alarms = { a[2] => alarm }
          sxl["objects"][a[0]]["alarms"] = alarms
        end
      else
        STDERR.puts "Object #{a[0]} not found"
      end

      y = y + 1
    end

  when "Aggregated status", "Aggregerad status" # swedish fix
    # Get the state bits
    y = 16
    state = {}
    for i in 0..7
      if sheet[y+i][2] and sheet[y+i][2].value
        state[i+1] = { 'title' => sheet[y+i][1].value,
                      'description' => sheet[y+i][2].value }
      else
        state[i+1] = { 'title' => sheet[y+i][1].value }
      end
    end

    y = 6
    while(sheet[y][0] != nil and sheet[y][0].value != nil) do


      # Get the basic aggregated status info
      agg = [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value,
              sheet[y][3].value, sheet[y][4].value ]

      # Add to yaml
      if sxl["objects"][agg[0]]
        sxl["objects"][agg[0]]["aggregated_status"] = state
      else
        STDERR.puts "Object #{a[0]} not found"
      end

      if options[:extended]
        sxl["objects"][agg[0]]["functional_position"] = agg[2]
        sxl["objects"][agg[0]]["functional_state"] = agg[3]
      end
      
      y = y + 1

    end
  when "Status"
    y = 6
    while(sheet[y][0] != nil and sheet[y][0].value != nil) do
      # Get the basic status info
      s = [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value,
            sheet[y][3].value ]

      # Get each argument
      x = 4
      a = {}
      while(sheet[y][x] != nil and sheet[y][x].value != nil and !sheet[y][x].value.empty?) do
        a[sheet[y][x].value] = {
            'type' => sheet[y][x+1].value,
             'description' => sheet[y][x+3].value
        }

        # No need to output values if type is boolean
        unless a[sheet[y][x].value]['type'] == 'boolean'
          # Output values in a different way
          if sheet[y][x+2].value.start_with?("-")
            # Values consists of several options
            values = sheet[y][x+2].value.split("-")
            values.shift

            values.each {|v|
              v.delete!("\n")

              # Try to find the corresponding description in
              # the description field using "key: value" format
              # Removes the key in description on match
              field = sheet[y][x+3].value
              desc = get_value(field, v)
              if desc.nil? then
                desc = ''
              end

              # Add to yaml
              v = to_integer(v)
              if a[sheet[y][x].value]['values']
                a[sheet[y][x].value]['values'][v] = desc
              else
                a[sheet[y][x].value]['values'] = {v => desc}
              end
            }
          else
            if options[:extended]
              a[sheet[y][x].value]['range'] = sheet[y][x+2].value
            end
          end
        end

        # Remove the description field if it is empty
        if a[sheet[y][x].value]['description'].empty?
          a[sheet[y][x].value].delete('description')
        end

        x = x + 4
      end
      status = {
        'description' => s[3],
        'arguments' => a
      }
      status.store("object", s[1]) if s[1] != nil

      # Add to yaml
      if sxl["objects"][s[0]]
        if sxl["objects"][s[0]]["statuses"]
          sxl["objects"][s[0]]["statuses"][s[2]] = status
        else
          statuses = { s[2] => status }
          sxl["objects"][s[0]]["statuses"] = statuses
        end
      else
        STDERR.puts "Object #{s[0]} not found"
      end

      y = y + 1
    end
  when "Commands"
    get_command_section(sheet).each { |y|
      while(sheet[y][0] != nil and sheet[y][0].value != nil) do
        # Get the basic command info
        c = [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value,
            sheet[y][3].value ]

        # Get each argument
        x = 4
        a = {}
        while(sheet[y][x] != nil and sheet[y][x].value != nil) do
          a[sheet[y][x].value] = {
            'type' => sheet[y][x+2].value,
            'description' => sheet[y][x+4].value
          }

          # No need to output values if type is boolean
          unless a[sheet[y][x].value]['type'] == 'boolean'
            # Output values in a different way
            if sheet[y][x+3].value.start_with?("-")
              values = sheet[y][x+3].value.split("-")
              values.shift
              values.each {|v|
                v.delete!("\n")

                # Try to find the corresponding description in
                # the description field using "key: value" format
                # Removes the key in description on match
                field = sheet[y][x+4].value
                desc = get_value(field, v)
                if desc.nil? then
                  desc = ''
                end

                # Add to yaml
                v = to_integer(v)
                if a[sheet[y][x].value]['values']
                  a[sheet[y][x].value]['values'][v.to_s] = desc
                else
                  a[sheet[y][x].value]['values'] = {v.to_s => desc}
                end
              }
            else
              if options[:extended]
                a[sheet[y][x].value]['range'] = sheet[y][x+3].value
              end
            end
          end

          # Remove the description field if it is empty
          if a[sheet[y][x].value]['description'].empty?
            a[sheet[y][x].value].delete('description')
          end


          co = sheet[y][x+1].value # command
          x = x + 5
        end

        command = {
          'description' => c[3],
          'arguments' => a,
          'command' => co
        }
        command.store("object", c[1]) if c[1] != nil

        # Add to yaml
        if sxl["objects"][c[0]]
          if sxl["objects"][c[0]]["commands"]
            sxl["objects"][c[0]]["commands"][c[2]] = command
          else
            commands = { c[2] => command }
            sxl["objects"][c[0]]["commands"] = commands
          end
        else
          STDERR.puts "Object #{c[0]} not found"
        end

        y = y + 1
      end
    }

  else
    # Objects sheets

    # site id
    siteid = sheet[1][1].value
    siteid_desc = sheet[1][2].value
    
    # grouped objects
    objects = get_object_section(sheet, 6, options)

    # single objects
    objects = objects.merge(get_object_section(sheet, 24, options))

    site = {
      'description' => siteid_desc,
      'objects' => objects
    }

    STDERR.puts "Warning: " + siteid + " already defined" if sites[siteid]
    sites[siteid] = site
  end
end

sxl["sites"] = sites if options[:site]
reindent(sxl.to_yaml)

