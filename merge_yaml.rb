#!/usr/bin/env ruby
require 'yaml'
require 'optparse'

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
usage = "Usage: mergeyaml.rb --objects [objects.yaml] --site [site.yaml]"
OptionParser.new do |opts|
  opts.banner = usage

  opts.on("--sxl [YAML]", "Signal Exchange List") do |x|
    options[:sxl] = x
  end

  opts.on("--site [YAML]", "Site configuration") do |t|
    options[:site] = t
  end
end.parse!

abort("--sxl needs to be set") if options[:sxl].nil?
abort("--site needs to be set") if options[:site].nil?

# Read yaml
objects = YAML.load_file(options[:sxl])
site = YAML.load_file(options[:site])

# Merge
#objects["sites"] = site["sites"]
site["objects"] = objects["objects"]

#reindent(objects.to_yaml)
reindent(site.to_yaml)
