#!/usr/bin/env ruby
require 'yaml'
require 'rubyXL'
require 'rubyXL/convenience_methods'
require 'optparse'

# yaml2xlsx reads sxl in yaml-format and outputs to xlsx-format

# Excel template defaults
@return_value_no=2
@return_value_start_col=4

def set_cell(sheet, col, row, string)
  return if string == nil
  col -= 1
  row -= 1
  if sheet[row] == nil or sheet[row][col] == nil
    sheet.add_cell row, col, string
  else
    sheet[row][col].change_contents string
  end
end

# Find row where first column is equal to string
def find_row(sheet, string)
  col = 0
  row = 0
  until row == sheet.sheet_data.rows.size
    unless sheet[row] == nil or sheet[row][col] == nil
      return row if sheet[row][col].value == string
    end
    row += 1
  end
  p "Could not find row"
end

options = {}
usage = "Usage: yaml2xlsx.rb [OPTIONS]--template <XLSX>] [sxl.yaml] [site.yaml]"
OptionParser.new do |opts|
  opts.banner = usage

  opts.on("--template [XLSX]", "Excel SXL Template") do |p|
    options[:template] = p
  end

  opts.on("--objects [YAML]", "Signal Exchange List (objects)") do |x|
    options[:objects] = x
  end

  opts.on("--site [YAML]", "Signal Exchange List (site)") do |t|
    options[:site] = t
  end

  opts.on("--short-desc", "Short description") do |s|
    options[:short] = s
  end

  opts.on("--stdout", "Output xlsx on STDOUT") do |o|
    options[:stdout] = o
  end
end.parse!

abort("--template needs to be set") if options[:template].nil?
abort("--objects needs to be set") if options[:objects].nil?
abort("--site needs to be set") if options[:site].nil?

workbook = RubyXL::Parser.parse(options[:template])

# Read yaml
objects = YAML.load_file(options[:objects])
site = YAML.load_file(options[:site])

# Version
sheet = workbook['Version']
set_cell(sheet, 2, 4, site["id"])            # Plant id
set_cell(sheet, 2, 6, site["description"])   # Plant name
set_cell(sheet, 2, 10, site["constructor"])  # Constructor
set_cell(sheet, 2, 12, site["reviewed"])     # Reviewed
set_cell(sheet, 2, 15, site["approved"])     # Approved
set_cell(sheet, 2, 18, site["created-date"]) # Created date
set_cell(sheet, 2, 21, site["version"])      # SXL revision number
set_cell(sheet, 3, 21, site["date"])         # SXL revision date
set_cell(sheet, 2, 26, site["rsmp-version"]) # RSMP version

# Object types
sheet = workbook['Object types']
gy = find_row(sheet, "Grouped object types")+3
sy = find_row(sheet, "Single object types")+3
objects["objects"].each { |object|
  # Is it a grouped object or not
  unless object[1]["aggregated_status"].nil?
    set_cell(sheet, 1, gy, object[0])
    set_cell(sheet, 2, gy, object[1]["description"])
    gy = gy + 1
  else
    set_cell(sheet, 1, sy, object[0])
    set_cell(sheet, 2, sy, object[1]["description"])
    sy = sy + 1
  end
}

# Objects
sheet = workbook['Objects']
gy = find_row(sheet, "Grouped objects")+3
sy = find_row(sheet, "Single objects")+3
site["sites"].each { |site|
  set_cell(sheet, 2, 2, site[0])
  set_cell(sheet, 3, 2, site[1]["description"])

  site[1]["objects"].each { |object|
    # Is it a grouped object or not
    unless object[1]["aggregated_status"].nil?

      object[1].each { |grouped|
        set_cell(sheet, 1, gy, object[0])
        set_cell(sheet, 2, gy, grouped[0])
        if(grouped[1])
          set_cell(sheet, 3, gy, grouped[1]["componentId"])
          set_cell(sheet, 4, gy, grouped[1]["ntsObjectId"])
          set_cell(sheet, 5, gy, grouped[1]["externalNtsId"])
          set_cell(sheet, 6, gy, grouped[1]["description"])
        else
          STDERR.puts "Warning! componentId is missing"
        end
        gy = gy + 1
      }
    else
      object[1].each { |single|
        set_cell(sheet, 1, sy, object[0])
        set_cell(sheet, 2, sy, single[0])
        if(single[1])
          set_cell(sheet, 3, sy, single[1]["componentId"])
          set_cell(sheet, 4, sy, single[1]["ntsObjectId"])
          set_cell(sheet, 5, sy, single[1]["externalNtsId"])
          set_cell(sheet, 6, sy, single[1]["description"])
        else
          STDERR.puts "Warning! componentId is missing"
        end
        sy = sy + 1
      }
    end
  }
}

# Aggregated status
sheet = workbook['Aggregated status']
y = 7
objects["objects"].each { |object|
  unless object[1]["aggregated_status"].nil?
    set_cell(sheet, 1, y, object[0])
    set_cell(sheet, 3, y, object[1]["functional_position"])
    set_cell(sheet, 4, y, object[1]["functional_state"])
    set_cell(sheet, 5, y, object[1]["aggregated_status_description"])
    set_cell(sheet, 3, 17, object[1]["aggregated_status"][1]["description"])
    set_cell(sheet, 3, 18, object[1]["aggregated_status"][2]["description"])
    set_cell(sheet, 3, 19, object[1]["aggregated_status"][3]["description"])
    set_cell(sheet, 3, 20, object[1]["aggregated_status"][4]["description"])
    set_cell(sheet, 3, 21, object[1]["aggregated_status"][5]["description"])
    set_cell(sheet, 3, 22, object[1]["aggregated_status"][6]["description"])
    set_cell(sheet, 3, 23, object[1]["aggregated_status"][7]["description"])
    set_cell(sheet, 3, 24, object[1]["aggregated_status"][8]["description"])
  end
}

# Alarm
alarm = Struct.new(:object_type, :object, :aid, :description,
                   :external_alarmcodeid, :external_ntsalarmcodeid, :priority, :category, :return_value)
return_value = Struct.new(:name, :type, :value, :comment)
a = []
r_list = []

# for each object type in yaml, look at all the alarms
objects["objects"].each { |object|
  if object[1]["alarms"]
    object[1]["alarms"].each { |item|
  
      # Return values
      r_list = []
      unless item[1]["arguments"].nil?
        item[1]["arguments"].each { |argument, value|
          r = []
          # Remove _list from the type (integer_list)
          value["type"].gsub!("_list", "")
  
          if value["type"] == "boolean"
              values = "-False\n-True"
          elsif value["type"] == "base64"
              values = "[base64]"
          else
            description = ""
            if not value["values"].nil?
              # Make a list of values and append to description
              values = ""
              value["values"].each { |v, desc |
                values += "-" + v.to_s + "\n"
                unless desc.empty?
                  description += + v.to_s + ": " + desc + "\n"
                end
              }
            elsif value["type"] == "string"
              values = "[string]"
            elsif not value["min"].nil?
              min = value["min"]
              max = value["max"]
              values = "[" + min.to_s + "-" + max.to_s + "]"
            else
              values = ""
            end
            values.chomp!
            description.chomp!
            if value["description"].nil?
              value["description"] = description
            else
              value["description"].concat("\n" + description)
              value["description"].chomp!
            end
          end
          r << return_value.new(argument, value["type"], values, value["description"])
          r_list.push(r)
        }
      end
    a << alarm.new(object[0], item[1]["object"], item[0], item[1]["description"], 
                  item[1]["externalAlarmCodeId"], item[1]["externalNtsAlarmCodeId"],
                  item[1]["priority"], item[1]["category"], r_list)

    }
  end
}

sheet = workbook['Alarms']
row = 7

# Sort by alarmId
a.sort_by { |ao| ao.aid }
a.each { |ao|
  set_cell(sheet, 1, row, ao.object_type)
  set_cell(sheet, 2, row, ao.object)
  set_cell(sheet, 3, row, ao.aid)
  if options[:short].nil?
    set_cell(sheet, 4, row, ao.description)
  else
    set_cell(sheet, 4, row, ao.description.lines.first.chomp)
  end
  set_cell(sheet, 5, row, ao.external_alarmcodeid)
  set_cell(sheet, 6, row, ao.external_ntsalarmcodeid)
  set_cell(sheet, 7, row, ao.priority)
  set_cell(sheet, 8, row, ao.category)

  # Return values
  col = 9
  unless ao.return_value[0].nil?
    ao.return_value[0].each { |rv|
      set_cell(sheet, col, row, rv.name)
      set_cell(sheet, col+1, row, rv.type)
      set_cell(sheet, col+2, row, rv.value)
      set_cell(sheet, col+3, row, rv.comment)
      col += 4
    } 
  end
  row += 1
}


# Status
status = Struct.new(:object_type, :object, :sid, :description, :return_value)
return_value = Struct.new(:name, :type, :value, :comment)
s = []
r_list = []

# for each object type in yaml, look at all the statuses
objects["objects"].each { |object|
  if object[1]["statuses"]
    object[1]["statuses"].each { |item|
  
      # Return values
      r_list = []
      unless item[1]["arguments"].nil?
        item[1]["arguments"].each { |argument, value|  # in statuses, it's called return value
          r = []
          # Remove _list from the type (integer_list)
          value["type"].gsub!("_list", "")
  
          if value["type"] == "boolean"
              values = "-False\n-True"
          elsif value["type"] == "base64"
              values = "[base64]"
          else
            description = ""
            if not value["values"].nil?
              # Make a list of values and append to description
              values = ""
              value["values"].each { |v, desc |
                values += "-" + v.to_s + "\n"
                unless desc.empty?
                  description += + v.to_s + ": " + desc + "\n"
                end
              }
            elsif value["type"] == "string"
              values = "[string]"
            elsif not value["min"].nil?
              min = value["min"]
              max = value["max"]
              values = "[" + min.to_s + "-" + max.to_s + "]"
            else
              values = ""
            end
            values.chomp!
            description.chomp!
            if value["description"].nil?
              value["description"] = description
            else
              value["description"].concat("\n" + description)
              value["description"].chomp!
            end
          end
          r << return_value.new(argument, value["type"], values, value["description"])
          r_list.push(r)
        }
      end
    s << status.new(object[0], item[1]["object"], item[0], item[1]["description"], r_list)
    }
  end
}

sheet = workbook['Status']
row = 7

# Sort by statusId
s.sort_by { |so| so.sid }
s.each { |so|
  set_cell(sheet, 1, row, so.object_type)
  set_cell(sheet, 2, row, so.object)
  set_cell(sheet, 3, row, so.sid)
  if options[:short].nil?
    set_cell(sheet, 4, row, so.description)
  else
    set_cell(sheet, 4, row, so.description.lines.first.chomp)
  end

  # Return values
  col = 5
  so.return_value[0].each { |rv|
    set_cell(sheet, col, row, rv.name)
    set_cell(sheet, col+1, row, rv.type)
    set_cell(sheet, col+2, row, rv.value)
    set_cell(sheet, col+3, row, rv.comment)
    col += 4
  }
  row += 1
}


# Commands
command = Struct.new(:object_type, :object, :cid, :description, :argument)
arg = Struct.new(:name, :command, :type, :value, :comment)
c = []
a_list = []

# for each object type in yaml, look at all the commands
objects["objects"].each { |object|
  if object[1]["commands"]
    object[1]["commands"].each { |item|
   
      # Arguments
      a_list = []
      item[1]["arguments"].each { |argument, value|
        a = []
        # Remove _list from the type (integer_list)
        value["type"].gsub!("_list", "")
  
        if value["type"] == "boolean"
            values = "-False\n-True"
        elsif value["type"] == "base64"
            values = "[base64]"
        else
          description = ""
          if not value["values"].nil?
            # Make a list of values and append to description
            values = ""
            value["values"].each { |v, desc |
              values += "-" + v.to_s + "\n"
              unless desc.empty?
                description += + v.to_s + ": " + desc + "\n"
              end
            }
          elsif value["type"] == "string"
            values = "[string]"
          elsif not value["min"].nil?
            min = value["min"]
            max = value["max"]
            values = "[" + min.to_s + "-" + max.to_s + "]"
          else
            values = ""
          end
          values.chomp!
          description.chomp!
          if value["description"].nil?
            value["description"] = description
          else
            value["description"].concat("\n" + description)
          end
        end
        a << arg.new(argument, item[1]["command"], value["type"], values, value["description"])
        a_list.push(a)
      }
      c << command.new(object[0], item[1]["object"], item[0], item[1]["description"], a_list)
    }
  end
}

# When converting from yaml to excel, put all commands under "parameter"
# since it is not possible to differentiate them further
sheet = workbook['Commands']
row = find_row(sheet, "Parameter")+3

# Sort by commandId
c.sort_by! { |co| co.cid }
c.each { |co|
  set_cell(sheet, 1, row, co.object_type)
  set_cell(sheet, 2, row, co.object)
  set_cell(sheet, 3, row, co.cid)
  if options[:short].nil?
    set_cell(sheet, 4, row, co.description)
  else
    set_cell(sheet, 4, row, co.description.lines.first.chomp)
  end

  # Arguments
  col = 5
  co.argument[0].each { |ar|
    set_cell(sheet, col, row, ar.name)
    set_cell(sheet, col+1, row, ar.command)
    set_cell(sheet, col+2, row, ar.type)
    set_cell(sheet, col+3, row, ar.value)
    set_cell(sheet, col+4, row, ar.comment)
    col += 5
  }
  row += 1
}

# Save
if options[:stdout].nil?
  workbook.write("output.xlsx")
else
  # write to stdout instead
  print workbook.stream.string
end
