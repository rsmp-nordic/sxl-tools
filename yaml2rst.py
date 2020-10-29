#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import argparse
import yaml
import re
from tabulate import tabulate

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

def start_figtable(widths,label):
    print("")
    print(".. figtable::")
    print("   :nofig:")
    print("   :label:", label)
    print("   :caption:", label)
    print("   :loc: H")
    print("   :spec: >{\\raggedright\\arraybackslash}", end='')
    sep = False
    for width in widths:
        if sep == True:
            print(" ", end='')
        sep = True
        print("p{", width, "\\linewidth}", sep='', end='')
    print("")
    print("")

def end_figtable():
    print("..")

def print_version():
    print("Signal Exchange List")
    print("====================")
    if "id" in yaml_sxl:
        print("+ **Plant Id**: "   + yaml_sxl['id'])
    if "description" in yaml_sxl:
        print("+ **Plant Name**: " + yaml_sxl['description'])
    if "constructor" in yaml_sxl:
        print("+ **Constructor**: " + yaml_sxl['constructor'])
    if "reviewed" in yaml_sxl:
        print("+ **Reviewed**: " + yaml_sxl['reviewed'])
    if "approved" in yaml_sxl:
        print("+ **Approved**: " + yaml_sxl['approved'])
    if "created-date" in yaml_sxl:
        print("+ **Created date**: " + yaml_sxl['created-date'])
    if "version-date" in yaml_sxl:
        print("+ **SXL revision**: " + yaml_sxl['version'])
    if "date" in yaml_sxl:
        print("+ **Revision date**: " + yaml_sxl['date'])
    if "rsmp_version" in yaml_sxl:
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
    start_figtable(widths, "Grouped objects")
    grouped = []
    grouped.append(table_headers)
    # For each object
    for object_name,object in yaml_sxl['objects'].items():
        if "aggregated_status" in object:
            grouped.append([object_name, object['description']])
    for line in tabulate(grouped, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    end_figtable()

    print("")
    print("Single objects")
    print("^^^^^^^^^^^^^^")
    widths = ["0.30", "0.50"]
    table_headers = ["ObjectType", "Description"]
    start_figtable(widths, "Single objects")
    single = []
    single.append(table_headers)
    for object_name,object in yaml_sxl['objects'].items():
        if "aggregated_status" not in object:
            single.append([object_name, object['description']])
    for line in tabulate(single, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    end_figtable()

def print_aggregated_status():
    print("")
    print("Aggregated status")
    print("-----------------")

    widths = ["0.15", "0.20", "0.18", "0.18", "0.15"]
    table_headers = ["ObjectType","Status","functionalPosition","functionalState","Description"]
    start_figtable(widths, "Aggregated status")
    agg_status = []
    agg_status.append(table_headers)
    # For each object
    for object_name,object in yaml_sxl['objects'].items():
        if "aggregated_status" in object:
            agg_status.append([object_name, "See state-bit definitions below", "", "", object['description']])

    for line in tabulate(agg_status, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    end_figtable()


    widths = ["0.15", "0.30", "0.45"]
    table_headers = ["State- Bit nr (12345678)", "Description", "Comment"]
    start_figtable(widths, "State bits")
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
    end_figtable()


def print_alarms():
    print("")
    print("Alarms")
    print("------")

    widths = ["0.15", "0.10", "0.45", "0.07", "0.07"]
    table_headers = ["ObjectType","alarmCodeId","Description","Priority","Category"]
    start_figtable(widths, "Alarms")

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
    end_figtable()

    # Print detailed alarm info
    # incl. return values
    alarms.sort(key=sort_cid)
    for object_name,alarm_id,description,priority,category in alarms:

        print("")
        print(alarm_id)
        print("^^^^^")
        print("")

        print(description.replace("\n", " |br| "))
        print("")

        return_values = []
        for object_name,object in yaml_sxl['objects'].items():
            for id,alarm in object['alarms'].items():
                if(alarm_id == id):
                    if "arguments" in alarm:
                        for argument_name, argument in alarm['arguments'].items():
                            name = argument_name
                            type = argument['type']
                            if "values" in argument:
                                val_list = []
                                for v in argument['values']:
                                    val_list.append("-" + v)
                                value = " |br| ".join(val_list)

                            else:
                                # Extended
                                if "range" in argument:
                                    value = argument['range']
                                else:
                                    if(type == "boolean"):
                                        value = "-True |br| -False"
                                    else:
                                        value = ""
                            if "description" in argument:
                                comment = argument['description'].replace("\n", " |br| ")
                            else:
                                comment = ""

                            # Add the full description in the comment
                            if "values" in argument:
                                for n,desc in argument['values'].items():
                                    if desc:
                                        comment += " |br| " + n + ": " + desc

                            return_values.append([name, type, value, comment])

        if return_values:
            widths = ["0.15", "0.08", "0.13", "0.35"]
            table_headers = ["Name","Type","Value","Comment"]
            start_figtable(widths, alarm_id)

            return_values.insert(0, table_headers)
            for line in tabulate(return_values, headers="firstrow", tablefmt="rst").splitlines():
                print('   ' + line)
            end_figtable()

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
    start_figtable(widths, "Status")

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
    end_figtable()

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
                    print(status['description'].replace("\n", " |br| "))
                    print("")

        widths = ["0.15", "0.08", "0.13", "0.50"]
        table_headers = ["Name", "Type", "Value", "Comment"]
        start_figtable(widths, status_id)

        return_values = []
        return_values.append(table_headers)
        for object_name,object in yaml_sxl['objects'].items():
            for id,status in object['statuses'].items():
                if(status_id == id):
                    if "arguments" in status:
                        for argument_name,argument in status['arguments'].items():
                            name = argument_name
                            type = argument['type']
                            if "values" in argument:
                                val_list = []
                                for v in argument['values']:
                                    val_list.append("-" + str(v))
                                value = " |br| ".join(val_list)
                            else:
                                # Extended
                                if "range" in argument:
                                    value = argument['range']
                                else:
                                    if(type == "boolean"):
                                        value = "-False |br| -True"
                                    else:
                                        value = ""
                            if "description" in argument:
                                comment = argument['description'].rstrip("\n")
                                comment = comment.replace("\n", " |br| ")
                            else:
                                comment = ""

                            # Add the full description in the comment
                            if "values" in argument:
                                for n,desc in argument['values'].items():
                                    if desc:
                                        if comment != "":
                                            comment += " |br| "
                                        comment += str(n) + ": " + str(desc)

                            return_values.append([name, type, value, comment])

        for line in tabulate(return_values, headers="firstrow", tablefmt="rst").splitlines():
            print('   ' + line)
        end_figtable()

def print_commands():
    print("")
    print("Commands")
    print("--------")

    widths = ["0.24", "0.15", "0.40"]
    table_headers = ["ObjectType","commandCodeId","Description"]
    start_figtable(widths, "Commands")

    command_table = []
    commands = []
    # For each object
    for object_name,object, in yaml_sxl['objects'].items():
        for command_id,command in object['commands'].items():
            command_table.append([object_name, '`' + command_id + '`_', command['description'].splitlines()[0]])
            commands.append([object_name, command_id, command['description'].replace("\n", " |br| ")])

    # Print command table
    # Sort and insert headers
    command_table.sort(key=sort_cid)
    command_table.insert(0, table_headers)
    for line in tabulate(command_table, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    end_figtable()

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
                    print(command['description'].replace("\n", " |br| "))
                    print("")

        widths = ["0.14", "0.20", "0.07", "0.15", "0.30"]
        table_headers = ["Name", "Command", "Type", "Value", "Comment"]
        start_figtable(widths, command_id)

        arguments = []
        for object_name,object in yaml_sxl['objects'].items():
            for id,command in object['commands'].items():
                if(command_id == id):
                    if "arguments" in command:
                        for argument_name,argument in command['arguments'].items():
                            name = argument_name
                            type = argument['type']
                            if "values" in argument:
                                val_list = []
                                for v in argument['values']:
                                    val_list.append("-" + v)
                                value = " |br| ".join(val_list)
                            else:
                                # Extended
                                if "range" in argument:
                                    value = argument['range']
                                else:
                                    if(type == "boolean"):
                                        value = "-False |br| -True"
                                    else:
                                        value = ""
                            if "description" in argument:
                                comment = argument['description'].rstrip("\n")
                                comment = comment.replace("\n", " |br| ")
                            else:
                                comment = ""

                            # Add the full description in the comment
                            if "values" in argument:
                                for n,desc in argument['values'].items():
                                    if desc:
                                        if comment != "":
                                            comment += " |br| "
                                        comment += str(n) + ": " + str(desc)

                            arguments.append([name, command['command'], type, value, comment])

        arguments.insert(0, table_headers)

        for line in tabulate(arguments, headers="firstrow", tablefmt="rst").splitlines():
            print('   ' + line)
        end_figtable()


# Read the yaml from stdin
yaml_sxl = yaml.safe_load(sys.stdin.read())

print_version()
print_object_types()
print_aggregated_status()
print_alarms()
print_status()
print_commands()
rst_line_break_substitution()
