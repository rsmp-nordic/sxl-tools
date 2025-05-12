#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import argparse
import pypandoc
import yaml
from yaml.resolver import Resolver
import re
from tabulate import tabulate

# Prevent PyYAML from converting On/Off/Yes/No into True/False
# remove resolver entries for On/Off/Yes/No
for ch in "OoYyNn":
    if len(Resolver.yaml_implicit_resolvers[ch]) == 1:
        del Resolver.yaml_implicit_resolvers[ch]
    else:
        Resolver.yaml_implicit_resolvers[ch] = [x for x in
                Resolver.yaml_implicit_resolvers[ch] if x[0] != 'tag:yaml.org,2002:bool']

def rst_line_break_substitution():
    print("");
    print(".. |br| replace:: |br_html| |br_latex|")
    print("")
    print(".. |br_html| raw:: html")
    print("")
    print("   <br>")
    print("")
    print(".. |br_latex| raw:: latex")
    print("")
    print("   \\newline")
    print("")

def sort_cid(alarm):
    return alarm[1].translate({ord(i): None for i in 'ASM\\`_'})

# Process description of alarm, status and command
# (description of each attribute/return value treated separately)
# - Removes last dot of first line
# - Inserts blank second line
# - Converts any markdown to RST
def trim_description(description):
    return md2rst(add_blank(rm_dot(description)))

# Convert markdown to restructuredText
def md2rst(description):
    return pypandoc.convert_text(description, 'rst', format='md')

# Removes trailing "." on first line
def rm_dot(description):
    desc = []
    desc = description.split("\n")

    # Strip "."
    if len(desc) > 0:
        if desc[0].endswith("."):
            desc[0] = desc[0].rstrip(".")
    return '\n'.join(desc)

# Adds a empty line after first line
def add_blank(description):
    desc = []
    desc = description.split("\n")

    # If second line is not empty, and line
    if len(desc) > 1 and desc[1] != "":
        desc.insert(1, "")

    return '\n'.join(desc)

def read_return_value(name, argument, reserved):
    arg_type = argument['type']

    array = []
    enum = {}
    min = ""
    max = ""

    # If the 'values' exists, use it to construct a dictionary
    if "values" in argument:
        if type(argument['values']) is dict:
            for n,desc in argument['values'].items():
                enum[n] = desc;
    else:
        if arg_type == "array":
            if "items" in argument:
                for arg_name, arg in argument['items'].items():
                    array.append(read_return_value(arg_name, arg, reserved))
            value = ""
        elif arg_type == "integer" or arg_type == "long" or arg_type == "float" or arg_type == "integer_as_string":
            if "min" in argument:
                min = argument['min']

            if "max" in argument:
                max = argument['max']

    if "description" in argument:
        comment = argument['description'].rstrip("\n")

        # First line should not end with "."
        comment = rm_dot(comment)

        comment = comment.replace("\n", " |br|\n")

        # Lines should never start with whitespace
        comment = comment.replace("\n |br|", "\n|br|")

        if "optional" in argument:
            if argument['optional'] is True:
                comment = "(Optional) " + comment

        if "deprecated" in argument:
            if argument['deprecated'] is True:
                comment = "``Deprecated`` " + comment
    else:
        comment = ""

    if reserved is True:
        comment = "``Reserved``"

    return name, arg_type, min, max, enum, comment, array

# Remove unused columns when printing arguments/return values
def remove_unused_columns(arg):
    # Remove both max and min if max is unused
    # ('0' evaluates to False, which min often is, so only max is checked)

    widths = ["0.15", "0.10", "0.10", "0.10", "0.20", "0.35"]
    headers = ["Name", "Type", "Min", "Max", "Enum", "Comment"]

    max_unused = True
    enum_unused = True

    # Check all rows
    for col in arg:
        if col[4]:
            enum_unused = False
        if col[3]:
            max_unused = False

    # Remove any unused cols
    for col in arg:
        if enum_unused:
            del col[4]
        if max_unused:
            del col[3]
            del col[2]

    # Update header
    if enum_unused:
        del widths[4]
        del headers[4]
        # enum width is 0.20, so if not used, add "0.10" to Name and Comment to fill the page width
        widths[0] = str(float(widths[0]) + 0.10)
        widths[4] = str(float(widths[4]) + 0.10)
    if max_unused:
        del widths[3]
        del headers[3]
        del widths[2]
        del headers[2]
        # min and max are 0.10, so if not used add 0.20 to comment to fill the page width
        widths[-1] = str(float(widths[-1]) + 0.20)
    return widths, headers, arg

def start_table(widths,label):
    print("")
    print(".. tabularcolumns:: ", end='')
    for width in widths:
        print("|\\Yl{", width, "}", sep='', end='')
    print("|")
    print("")
    print(".. table:: " + label)
    print("   :class: longtable")
    print("")
    print("")

def print_version():
    print("Signal Exchange List")
    print("====================")
    if "id" in yaml_sxl and args.extended:
        print("+ **Plant Id**: "   + yaml_sxl['id'])
    if "description" in yaml_sxl and args.extended:
        print("+ **Plant Name**: " + yaml_sxl['description'])
    if "constructor" in yaml_sxl and args.extended:
        print("+ **Constructor**: " + yaml_sxl['constructor'])
    if "reviewed" in yaml_sxl and args.extended:
        print("+ **Reviewed**: " + yaml_sxl['reviewed'])
    if "approved" in yaml_sxl and args.extended:
        print("+ **Approved**: " + yaml_sxl['approved'])
    if "created-date" in yaml_sxl and args.extended:
        print("+ **Created date**: " + yaml_sxl['created-date'])
    if "version-date" in yaml_sxl and args.extended:
        print("+ **SXL revision**: " + yaml_sxl['version'])
    if "date" in yaml_sxl and args.extended:
        print("+ **Revision date**: " + yaml_sxl['date'])
    if "rsmp_version" in yaml_sxl and args.extended:
        print("+ **RSMP version**: " + yaml_sxl['rsmp-version'])

def print_object_types():
    print("")
    print("Object Types")
    print("------------")
    print("")

    print("Grouped objects")
    print("^^^^^^^^^^^^^^^")
    widths = ["0.30", "0.50"]
    table_headers = ["ObjectType", "Description"]
    start_table(widths, "Grouped objects")
    grouped = []
    grouped.append(table_headers)
    # For each object
    for object_name,object in yaml_sxl['objects'].items():
        if "aggregated_status" in object:
            grouped.append([object_name, object['description']])
    for line in tabulate(grouped, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    print("")

    print("")
    print("Single objects")
    print("^^^^^^^^^^^^^^")
    widths = ["0.30", "0.50"]
    table_headers = ["ObjectType", "Description"]
    start_table(widths, "Single objects")
    single = []
    single.append(table_headers)
    for object_name,object in yaml_sxl['objects'].items():
        if "aggregated_status" not in object:
            single.append([object_name, object['description']])
    for line in tabulate(single, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    print("")

def print_aggregated_status():
    print("")
    print("Aggregated status")
    print("-----------------")

    widths = ["0.20", "0.20", "0.20", "0.40"]
    table_headers = ["ObjectType","functionalPosition","functionalState","Description"]
    start_table(widths, "Aggregated status")
    agg_status = []
    agg_status.append(table_headers)
    # For each object
    for object_name,object in yaml_sxl['objects'].items():
        if "aggregated_status" in object:

            # Functional position
            fP = ""
            if "functional_position" in object and object['functional_position']:
                fP_list = []
                if type(object['functional_position']) is list:
                    for pos in object['functional_position']:
                        fP_list.append("-" + pos)
                fP = " |br|\n".join(fP_list)

            # Functional state
            fS = ""
            if "functional_state" in object and object['functional_state']:
                fS_list = []
                if type(object['functional_state']) is list:
                    for pos in object['functional_state']:
                        fS_list.append("-" + pos)
                fS = " |br|\n".join(fS_list)

            # Aggregated status description
            as_desc = ""
            if not object['functional_position']:
                as_desc = "functionalPosition not used (set to null)"
            if not object['functional_state']:
                as_desc += " |br|\nfunctionalState not used (set to null)"

            agg_status.append([object_name, fP, fS, as_desc])

    for line in tabulate(agg_status, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    print("")

    widths = ["0.10", "0.30", "0.60"]
    table_headers = ["State-Bit", "Description", "Comment"]
    start_table(widths, "State bits")
    state_bits = []
    state_bits.append(table_headers)
    # For each object
    for object_name,object in yaml_sxl['objects'].items():
        if "aggregated_status" in object:
            for state_id,state in object['aggregated_status'].items():
                if "description" in state:
                    state_bits.append([state_id, state['title'], state['description'].replace("\n", " ")])
                else:
                    state_bits.append([state_id, state['title'], ""])

    for line in tabulate(state_bits, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    print("")


def print_alarms():
    print("")
    print("Alarms")
    print("------")

    widths = ["0.20", "0.10", "0.50", "0.10", "0.10"]
    table_headers = ["ObjectType","alarmCodeId","Description","Priority","Category"]
    start_table(widths, "Alarms")

    alarm_table = []
    alarms = []
    # For each object
    for object_name,object in yaml_sxl['objects'].items():
        for alarm_id,alarm in object['alarms'].items():
            if "reserved" in alarm and alarm['reserved'] is True:
                alarm['description'] = "``Reserved``"
            desc = rm_dot(alarm['description'])
            alarm_table.append([object_name, '`' + alarm_id + '`_', desc.splitlines()[0], alarm['priority'], alarm['category']])
            alarms.append([object_name, alarm_id, desc, alarm['priority'], alarm['category']])

    # Print alarm table
    # Sort and insert headers
    alarm_table.sort(key=sort_cid)
    alarm_table.insert(0, table_headers)
    for line in tabulate(alarm_table, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    print("")

    # Print detailed alarm info
    # incl. return values
    alarms.sort(key=sort_cid)
    for object_name,alarm_id,description,priority,category in alarms:

        print("")
        print(alarm_id)
        print("^^^^^")
        print("")

        print(trim_description(description))
        print("")

        return_values = []
        for object_name,object in yaml_sxl['objects'].items():
            for id,alarm in object['alarms'].items():
                if(alarm_id == id):
                    reserved = False
                    if "reserved" in alarm and alarm["reserved"] is True:
                        reserved = True
                    if "arguments" in alarm:
                        for argument_name, argument in alarm['arguments'].items():
                            name, type, min, max, enum, comment, array = read_return_value(argument_name, argument, reserved)
                            return_values.append([name, type, min, max, enum, comment])

        if return_values:
            # Remove unused columns
            widths, table_headers, return_values = remove_unused_columns(return_values)

            start_table(widths, alarm_id)
            return_values.insert(0, table_headers)

            for line in tabulate(return_values, headers="firstrow", tablefmt="rst").splitlines():
                print('   ' + line)
            print("")

def print_status():
    print("")
    print("Status")
    print("------")

    print("")
    print(".. raw:: latex")
    print("")
    print("    \\newpage")
    print("")

    widths = ["0.30", "0.10", "0.60"]
    table_headers = ["ObjectType","statusCodeId","Description"]
    start_table(widths, "Status")

    status_table = []
    statuses = []
    # For each object
    for object_name,object, in yaml_sxl['objects'].items():
        for status_id,status in object['statuses'].items():
            if "reserved" in status and status['reserved'] is True:
                status['description'] = "``Reserved``"
            desc = rm_dot(status['description'])
            status_table.append([object_name, '`' + status_id + '`_', desc.splitlines()[0]]) 
            statuses.append([object_name, status_id, desc])

    # Print status table
    # Sort and insert headers
    status_table.sort(key=sort_cid)
    status_table.insert(0, table_headers)
    for line in tabulate(status_table, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    print("")

    # Print detailed status info
    # incl. return values
    statuses.sort(key=sort_cid)
    for object_name,status_id,description in statuses:
        print("")
        print(status_id)
        print("^^^^^^^^")
        print("")

        # Print status description
        for object_name,object in yaml_sxl['objects'].items():
            for id,status in object['statuses'].items():
                if(id == status_id):

                    # Don't print if reserved for future use
                    if "reserved" in status and status['reserved'] is True:
                        print("``Reserved``")
                    else:
                        print(trim_description(status['description']))
                    print("")

        return_values = []
        array_values = {}
        for object_name,object in yaml_sxl['objects'].items():
            for id,status in object['statuses'].items():
                if(status_id == id):
                    reserved = False
                    if "reserved" in status and status["reserved"] is True:
                        reserved = True
                    if "arguments" in status:
                        for argument_name,argument in status['arguments'].items():
                            name, type, min, max, enum, comment, array = read_return_value(argument_name, argument, reserved)
                            return_values.append([name, type, min, max, enum, comment])
                            if(type == "array"):
                                for a in array:
                                    if not argument_name in array_values:
                                        array_values[argument_name] = []
                                    array_values[argument_name].append([a[0], a[1], a[2], a[3], a[4], a[5]])

        if return_values:
            # Remove unused columns
            widths, table_headers, return_values = remove_unused_columns(return_values)

            start_table(widths, status_id)
            return_values.insert(0, table_headers)

            for line in tabulate(return_values, headers="firstrow", tablefmt="rst").splitlines():
                print('   ' + line)
            print("")

            for name in array_values:
                # Remove unused columns
                widths, table_headers, array_values = remove_unused_columns(array_values)
                start_table(widths, status_id + " " + name)

                array_values[name].insert(0, table_headers)
                for line in tabulate(array_values[name], headers="firstrow", tablefmt="rst").splitlines():
                    print('   ' + line)
                print("")

def print_commands():
    print("")
    print("Commands")
    print("--------")

    widths = ["0.30", "0.15", "0.20", "0.35"]
    table_headers = ["ObjectType","commandCodeId","Command","Description"]
    start_table(widths, "Commands")

    command_table = []
    commands = []
    # For each object
    for object_name,object, in yaml_sxl['objects'].items():
        for command_id,command in object['commands'].items():
            if "reserved" in command and command['reserved'] is True:
                command['description'] = "``Reserved``"
            desc = rm_dot(command['description'])
            command_table.append([object_name, '`' + command_id + '`_', command['command'], desc.splitlines()[0]])
            commands.append([object_name, command_id, desc.replace("\n", " |br| ")])

    # Print command table
    # Sort and insert headers
    command_table.sort(key=sort_cid)
    command_table.insert(0, table_headers)
    for line in tabulate(command_table, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    print("")

    # Arguments
    commands.sort(key=sort_cid)
    for object_name,command_id,description in commands:
        print("")
        print(command_id)
        print("^^^^^")
        print("")

        # Print command description
        for object_name,object in yaml_sxl['objects'].items():
            for id,command in object['commands'].items():
                if(id == command_id):

                    # Don't print if reserved for future use
                    if "reserved" in command and command['reserved'] is True:
                        print("``Reserved``")
                    else:
                        print(trim_description(command['description']))
                    print("")

        arguments = []
        for object_name,object in yaml_sxl['objects'].items():
            for id,command in object['commands'].items():
                if(command_id == id):
                    reserved = False
                    if "reserved" in command and command["reserved"] is True:
                        reserved = True
                    if "arguments" in command:
                        for argument_name,argument in command['arguments'].items():
                            name, type, min, max, enum, comment, array = read_return_value(argument_name, argument, reserved)
                            arguments.append([name, type, min, max, enum, comment])

        if arguments:
            # Remove unused columns
            widths, table_headers, arguments = remove_unused_columns(arguments)

            start_table(widths, command_id)
            arguments.insert(0, table_headers)

            for line in tabulate(arguments, headers="firstrow", tablefmt="rst").splitlines():
                print('   ' + line)
            print("")

parser = argparse.ArgumentParser(description='Convert SXL in yaml to rst format')
parser.add_argument('--extended', action=argparse.BooleanOptionalAction)
args = parser.parse_args()

# Read the yaml from stdin
yaml_sxl = yaml.safe_load(sys.stdin.read())

print_version()
print_object_types()
print_aggregated_status()
print_alarms()
print_status()
print_commands()
rst_line_break_substitution()
