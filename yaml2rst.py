#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import sys
import argparse
import yaml
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
    table_headers = ["State- Bit nr (1234567)", "Description", "Comment"]
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

    alarms = []
    alarms_with_return_values = []
    # For each object
    for object_name,object in yaml_sxl['objects'].items():
        for alarm_id,alarm in object['alarms'].items():
            if "arguments" in alarm:
                alarms.append([object_name, '`' + alarm_id + '`_', alarm['description'], alarm['priority'], alarm['category']])
                for argument_name, argument in alarm['arguments'].items():
                    if alarm_id not in alarms_with_return_values:
                        alarms_with_return_values.append(alarm_id)
            else:
                alarms.append([object_name, alarm_id, alarm['description'], alarm['priority'], alarm['category']])

    # Sort
    alarms.sort(key=sort_cid)
    alarms.insert(0, table_headers)

    for line in tabulate(alarms, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    end_figtable()

    # Return values
    for alarm_id in alarms_with_return_values:
        print("")
        print(alarm_id)
        print("^^^^^")
        print("")

        # Print alarm description
        for object_name,object in yaml_sxl['objects'].items():
            for id,alarm in object['alarms'].items():
                if(id == alarm_id):
                    print(alarm['description'].replace("\n", " |br| "))
                    print("")

        widths = ["0.15", "0.08", "0.13", "0.35"]
        table_headers = ["Name","Type","Value","Comment"]
        start_figtable(widths, alarm_id)

        return_values = []
        return_values.append(table_headers)
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
                                # Extended value
                                if "value" in argument:
                                    value = argument['value']
                                else:
                                    value = ""
                            if "description" in argument:
                                comment = argument['description'].replace("\n", " |br| ")
                            else:
                                comment = ""
                            return_values.append([name, type, value, comment])

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

    statuses = []
    statuses_with_return_values = []
    # For each object
    for object_name,object, in yaml_sxl['objects'].items():
        for status_id,status in object['statuses'].items():
            if "arguments" in status:
                statuses.append([object_name, '`' + status_id + '`_', status['description']])
                for argument_name, argument in status['arguments'].items():
                    if status_id not in statuses_with_return_values:
                        statuses_with_return_values.append(status_id)
            else:
                statuses.append([object_name, status_id, status['description']])

    # Sort
    statuses.sort(key=sort_cid)
    statuses.insert(0, table_headers)
    statuses_with_return_values.sort()

    for line in tabulate(statuses, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    end_figtable()

    # Return values
    for status_id in statuses_with_return_values:
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
                                    val_list.append("-" + v)
                                value = " |br| ".join(val_list)
                            else:
                                # Extended value
                                if "value" in argument:
                                    value = argument['value']
                                else:
                                    value = ""
                            if "description" in argument:
                                comment = argument['description'].rstrip("\n")
                                comment = comment.replace("\n", " |br| ")
                            else:
                                comment = ""
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

    commands = []
    commands_with_arguments = []
    # For each object
    for object_name,object, in yaml_sxl['objects'].items():
        for command_id,command in object['commands'].items():
            commands.append([object_name, '`' + command_id + '`_', command['description'].replace("\n", " |br| ")])
            if "arguments" in command:
                for argument_name, argument in command['arguments'].items():
                    if command_id not in commands_with_arguments:
                        commands_with_arguments.append(command_id)
            else:
                commands.append([object_name, command_id, command['description']])

    # Sort
    commands.sort(key=sort_cid)
    commands.insert(0, table_headers)
    commands_with_arguments.sort()

    for line in tabulate(commands, headers="firstrow", tablefmt="rst").splitlines():
        print('   ' + line)
    end_figtable()

    # Arguments
    for command_id in commands_with_arguments:
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
                                value = argument['values'].replace("\n", " |br| ")
                            else:
                                value = ""
                            if "description" in argument:
                                comment = argument['description'].rstrip("\n")
                                comment = comment.replace("\n", " |br| ")
                            else:
                                comment = ""
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
