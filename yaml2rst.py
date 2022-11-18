#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import argparse
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
    return alarm[1].translate({ord(i): None for i in 'ASM\`_'})

def read_return_value(name, argument):
    arg_type = argument['type'].replace("_list", "")

    array = []

    # If the 'values' exists, use it to construct a list
    if "values" in argument:
        val_list = []
        for v in argument['values']:
            val_list.append("-" + str(v))
        value = " |br|\n".join(val_list)

    else:
        if arg_type == "boolean":
            value = "-False |br|\n-True"
        elif arg_type == "string":
            value = "[string]"
        elif arg_type == "base64":
            value = "[base64]"
        elif arg_type == "array":
            if "items" in argument:
                for name, arg in argument['items'].items():
                    array.append(read_return_value(name, arg))
            value = ""
        elif arg_type == "integer" or arg_type == "long" or arg_type == "float":
            if "min" in argument:
                min = argument['min']
            else:
                min = ""
            if "max" in argument:
                max = argument['max']
            else:
                max = ""
            value = "[" + str(min) + "-" + str(max) + "]"
        else:
            value = ""
    if "description" in argument:
        comment = argument['description'].rstrip("\n")
        comment = comment.replace("\n", " |br|\n")

        # Lines should never start with whitespace
        comment = comment.replace("\n |br|", "\n|br|")
    else:
        comment = ""

    # Add the full description in the comment
    if "values" in argument:
        if type(argument['values']) is dict:
            for n,desc in argument['values'].items():
                if desc:
                    if comment != "":
                        comment += " |br|\n"
                    comment += str(n) + ": " + str(desc)

    return name, arg_type, value, comment, array

def start_table(widths,label):
    print("")
    print(".. tabularcolumns:: ", end='')
    for width in widths:
        print("|\Yl{", width, "}", sep='', end='')
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

    widths = ["0.15", "0.16", "0.16", "0.40"]
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
                if type(object['functional_position']) is dict:
                    for agg_name,agg in object['functional_position'].items():
                        fP_list.append("-" + agg_name)
                fP = " |br| ".join(fP_list)

            # Functional state
            fS = ""
            if "functional_state" in object and object['functional_state']:
                fS_list = []
                if type(object['functional_state']) is dict:
                    for agg_name,agg in object['functional_state'].items():
                        fS_list.append("-" + agg_name)
                fS = " |br| ".join(fS_list)

            # Aggregated status description
            as_desc = ""
            if "aggregated_status_description" in object:
                as_desc = object['aggregated_status_description']

            agg_status.append([object_name, fP, fS, as_desc])

    for line in tabulate(agg_status, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    print("")

    widths = ["0.10", "0.30", "0.50"]
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

    widths = ["0.15", "0.10", "0.45", "0.07", "0.07"]
    table_headers = ["ObjectType","alarmCodeId","Description","Priority","Category"]
    start_table(widths, "Alarms")

    alarm_table = []
    alarms = []
    # For each object
    for object_name,object in yaml_sxl['objects'].items():
        for alarm_id,alarm in object['alarms'].items():
            alarm_table.append([object_name, '`' + alarm_id + '`_', alarm['description'].splitlines()[0], alarm['priority'], alarm['category']])
            alarms.append([object_name, alarm_id, alarm['description'], alarm['priority'], alarm['category']])

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

        print(description)
        print("")

        return_values = []
        for object_name,object in yaml_sxl['objects'].items():
            for id,alarm in object['alarms'].items():
                if(alarm_id == id):
                    if "arguments" in alarm:
                        for argument_name, argument in alarm['arguments'].items():
                            name, type, value, comment, array = read_return_value(argument_name, argument)
                            return_values.append([name, type, value, comment])

        if return_values:
            widths = ["0.15", "0.15", "0.20", "0.35"]
            table_headers = ["Name","Type","Value","Comment"]
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

    widths = ["0.24", "0.10", "0.55"]
    table_headers = ["ObjectType","statusCodeId","Description"]
    start_table(widths, "Status")

    status_table = []
    statuses = []
    # For each object
    for object_name,object, in yaml_sxl['objects'].items():
        for status_id,status in object['statuses'].items():
            status_table.append([object_name, '`' + status_id + '`_', status['description'].splitlines()[0]])
            statuses.append([object_name, status_id, status['description']])

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
                    print(status['description'])
                    print("")

        return_values = []
        array_values = {}
        for object_name,object in yaml_sxl['objects'].items():
            for id,status in object['statuses'].items():
                if(status_id == id):
                    if "arguments" in status:
                        for argument_name,argument in status['arguments'].items():
                            name, type, value, comment, array = read_return_value(argument_name, argument)
                            return_values.append([name, type, value, comment])
                            if(type == "array"):
                                for a in array:
                                    array_values[argument_name].append([a[0], a[1], a[2], a[3]])

        if return_values:
            widths = ["0.15", "0.15", "0.20", "0.50"]
            table_headers = ["Name", "Type", "Value", "Comment"]
            start_table(widths, status_id)

            return_values.insert(0, table_headers)
            for line in tabulate(return_values, headers="firstrow", tablefmt="rst").splitlines():
                print('   ' + line)
            print("")

            for name in array_values:
                widths = ["0.15", "0.15", "0.20", "0.50"]
                table_headers = ["Name", "Type", "Value", "Comment"]
                start_table(widths, status_id + " " + name)

                array_values[name].insert(0, table_headers)
                for line in tabulate(array_values[name], headers="firstrow", tablefmt="rst").splitlines():
                    print('   ' + line)
                print("")

def print_commands():
    print("")
    print("Commands")
    print("--------")

    widths = ["0.24", "0.15", "0.21", "0.21"]
    table_headers = ["ObjectType","commandCodeId","Command","Description"]
    start_table(widths, "Commands")

    command_table = []
    commands = []
    # For each object
    for object_name,object, in yaml_sxl['objects'].items():
        for command_id,command in object['commands'].items():
            command_table.append([object_name, '`' + command_id + '`_', command['command'], command['description'].splitlines()[0]])
            commands.append([object_name, command_id, command['description'].replace("\n", " |br| ")])

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
                    print(command['description'])
                    print("")

        arguments = []
        for object_name,object in yaml_sxl['objects'].items():
            for id,command in object['commands'].items():
                if(command_id == id):
                    if "arguments" in command:
                        for argument_name,argument in command['arguments'].items():
                            name, type, value, comment, array = read_return_value(argument_name, argument)
                            arguments.append([name, type, value, comment])

        if arguments:
            widths = ["0.14",  "0.14", "0.20", "0.45"]
            table_headers = ["Name", "Type", "Value", "Comment"]
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
