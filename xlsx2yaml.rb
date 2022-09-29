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
    return [ sheet[y][0].value, nil, nil, nil, nil, nil ]
  end

  # Empty component id
  if sheet[y][2] == nil
    return [ sheet[y][0].value, sheet[y][1].value, nil, nil, nil, nil ]
  end

  # Empty NTSObjectId
  if sheet[y][3] == nil
    return [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value, nil, nil, nil ]
  end

  # Empty externalNtsId
  if sheet[y][4] == nil
    return [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value, sheet[y][3].value, nil, nil ]
  end

  # Empty description
  if sheet[y][5] == nil
    return [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value, sheet[y][3].value, sheet[y][4].value, nil ]
  end

  [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value, sheet[y][3].value, sheet[y][4].value, sheet[y][5].value]
end

# Get objects from a given section in the object sheet
def get_object_section(sheet, type, options)
  # type = 1 Grouped objects
  # type = 2 Single objects

  y = 4 # Section won't start before row 4
  while(y<200) do
    if sheet[y] != nil and sheet[y][0] != nil
      break if sheet[y][0].value.eql?("Grouped objects") and type == 1
      break if sheet[y][0].value.eql?("Grouped object types") and type == 1
      break if sheet[y][0].value.eql?("Single objects") and type == 2
      break if sheet[y][0].value.eql?("Single object types") and type == 2
    end
    y = y + 1
  end
  y = y + 2 # Begins two lines after title

  objects = {}
  while(object = get_object(sheet, y)) do
    key = object[0]  # e.g. Traffic Controller, Signal Group

    # Check if object type already exists, merge otherwise
    unless objects[key]
      # Object types has no componentId
      if object[2] == nil
        objects[key] = { 'description' => object[1] }
      else
        objects[key] = { object[1] => { 'componentId' => object[2], 'ntsObjectId' => object[3] } }
        objects[key][object[1]].store("externalNtsId", to_integer(object[4])) if object[4] != nil
        objects[key][object[1]].store("description", object[5]) if object[5] != nil
      end
    else
        newobject = { object[1] => { 'componentId' => object[2], 'ntsObjectId' => object[3] } }
        newobject[object[1]].store("externalNtsId", to_integer(object[4])) if object[4] != nil
        newobject[object[1]].store("description", object[5]) if object[5] != nil
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
  return if field == nil

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
  indent = ""
  text.each_line do |line|
    if line == "\n"
      # Check white spaces in the beginning of previous line
      indent = prev_line.match(/^([ ]+)/)[0] rescue false
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

  opts.on("-o", "--object", "Output object information") do |o|
    options[:object] = o
  end

  opts.on("-s", "--site", "Output site information") do |s|
    options[:site] = s
  end

end.parse!

if ARGV.length < 1
  # Read from STDIN instead
  buffer = STDIN.read
  workbook = RubyXL::Parser.parse_buffer(buffer)
else
  XLSX = ARGV[0]
  workbook = RubyXL::Parser.parse(XLSX)
end
sxl = Hash.new

sites = {}
workbook.each do |sheet|
  case sheet.sheet_name
  when "Version"
    if options[:site]
      sxl["id"] = sheet[3][1].value if sheet[3]
      sxl["version"] = sheet[20][1].value
      sxl["date"] = sheet[20][2].value
      sxl["description"] = sheet[5][1].value
      sxl["constructor"] = sheet[9][1].value if sheet[9]
      sxl["reviewed"] = sheet[11][1].value if sheet[11][1].value
      sxl["approved"] = sheet[13][1].value if sheet[13] and sheet[13][1]
      sxl["created-date"] = sheet[17][1].value if sheet[17]
      sxl["rsmp-version"] = sheet[25][1].value if sheet[25]
    end
  when "Object types"
    # grouped objects
    sxl["objects"] = get_object_section(sheet, 1, options)

    # single objects
    sxl["objects"] = sxl["objects"].merge(get_object_section(sheet, 2, options))
  when "Alarms"
    y = 6
    while(sheet[y][0] != nil and sheet[y][0].value != nil) do
      # Get basic alarm info
      a = [ sheet[y][0].value, sheet[y][1].value, sheet[y][2].value,
            sheet[y][3].value, sheet[y][4].value, sheet[y][5].value,
            sheet[y][6].value, sheet[y][7].value] rescue false
      
      if a == false
        y = y + 1
        next
      end

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
              desc = get_value(rv[sheet[y][x].value]['description'], v)
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
            # Set 'max' and 'min' if type is integer, long or real
            if rv[sheet[y][x].value]['type'] == 'integer' or rv[sheet[y][x].value]['type'] == 'long' or rv[sheet[y][x].value]['type'] == 'real'
              values = sheet[y][x+2].value.tr('[]','').split("-")
              rv[sheet[y][x].value]['min'] = values[0].to_i
              rv[sheet[y][x].value]['max'] = values[1].to_i
            end
          end
        end

        # Remove the description field if it is empty
        if rv[sheet[y][x].value]['description'] and rv[sheet[y][x].value]['description'].empty?
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
        STDERR.puts "Object #{agg[0]} not found"
      end

      sxl["objects"][agg[0]]["functional_position"] = agg[2]
      sxl["objects"][agg[0]]["functional_state"] = agg[3]
      sxl["objects"][agg[0]]["aggregated_status_description"] = agg[4]
      
      y = y + 1

    end
  when "Status"
    y = 6
    while(sheet[y] and sheet[y][0] and sheet[y][0].value) do
      # Get the basic status info
      sheet[y][0] ? object_type = sheet[y][0].value : object_type = ''
      sheet[y][1] ? object = sheet[y][1].value : object = ''
      sheet[y][2] ? sci = sheet[y][2].value : sid = ''
      sheet[y][3] ? desc = sheet[y][3].value : desc = ''
      s = [ object_type, object, sci, desc ]

      # Get each argument
      x = 4
      a = {}
      while(sheet[y][x] != nil and sheet[y][x].value != nil and !sheet[y][x].value.empty?) do
        a[sheet[y][x].value] = {
            'type' => sheet[y][x+1].value,
             'description' => sheet[y][x+3].value.chomp
        }

        # No need to output values if type is boolean
        unless a[sheet[y][x].value]['type'] == 'boolean'
          # Output values in a different way
          if sheet[y][x+2].value and sheet[y][x+2].value.start_with?("-")
            # Values consists of several options
            values = sheet[y][x+2].value.split("-")
            values.shift

            values.each {|v|
              v.delete!("\n")

              # Try to find the corresponding description in
              # the description field using "key: value" format
              desc = get_value(a[sheet[y][x].value]['description'], v)
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
            # Set 'max' and 'min' if type is integer, long or real
            if a[sheet[y][x].value]['type'] == 'integer' or a[sheet[y][x].value]['type'] == 'long' or a[sheet[y][x].value]['type'] == 'real'
              values = sheet[y][x+2].value.tr('[]','').split("-")
              a[sheet[y][x].value]['min'] = values[0].to_i
              a[sheet[y][x].value]['max'] = values[1].to_i
              a[sheet[y][x].value]['type'] << "_list"	# Add _list to type if min/max is used
            end
          end
        end

        # Remove the description field if it is empty
        if a[sheet[y][x].value]['description'] and a[sheet[y][x].value]['description'].empty?
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
      while(sheet[y] != nil and sheet[y][0] != nil and sheet[y][0].value != nil) do
        # Get the basic command info
        sheet[y][0] ? object_type = sheet[y][0].value : object_type = ''
        sheet[y][1] ? object = sheet[y][1].value : object = ''
        sheet[y][2] ? cci = sheet[y][2].value : sid = ''
        sheet[y][3] ? desc = sheet[y][3].value : desc = ''
        c = [ object_type, object, cci, desc ]

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
                desc = get_value(a[sheet[y][x].value]['description'], v)
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
              # Set 'max' and 'min' if type is integer, long or real
              if a[sheet[y][x].value]['type'] == 'integer' or a[sheet[y][x].value]['type'] == 'long' or a[sheet[y][x].value]['type'] == 'real'
                values = sheet[y][x+3].value.tr('[]','').split("-")
                a[sheet[y][x].value]['min'] = values[0].to_i
                a[sheet[y][x].value]['max'] = values[1].to_i
                a[sheet[y][x].value]['type'] << "_list"	# Add _list to type if min/max is used
              end
            end
          end

          # Remove the description field if it is empty
          if a[sheet[y][x].value]['description'] and a[sheet[y][x].value]['description'].empty?
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
    objects = get_object_section(sheet, 1, options)

    # single objects
    objects = objects.merge(get_object_section(sheet, 2, options))

    site = {
      'description' => siteid_desc,
      'objects' => objects
    }

    STDERR.puts "Warning: " + siteid + " already defined" if sites[siteid]
    sites[siteid] = site
  end
end

sxl["sites"] = sites if options[:site]

# Clear "objects" if site should be used
sxl.delete("objects") if options[:site]

reindent(sxl.to_yaml)

