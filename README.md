sxl-tools
=========

Exports signal exchange lists (SXL) used by RSMP from Excel-format (xlsx) to
various formats.

* **sxl2csv.pl**   - Exports SXL from XLSX format to CSV files, zip compressed.
                     Suitable for use with the RSMP simulators
* **xlsx2yaml.rb** - Reads SXL in Excel format and outputs to YAML format
* **yaml2xlsx.rb** - Reads SXL in YAML format and outputs to Excel format
* **yaml2rst.py**  - Reads SXL in YAML format and outputs to RST format

Excel macro:
* **rsmp_clean.xlsx** - 1.0.7 SXL, adapt according to cId/siteid and num SG/DET

Notes about sxl2csv
-------------------

* Requires Spreadsheet::XLSX
* Ubuntu: `sudo apt install libspreadsheet-xlsx-perl`
* Arch Linux: `pacman -S zip` and from AUR: `perl-spreadsheet-xlsx` and
  dependencies

Notes about xlsx2yaml
---------------------

* Requires: gem install rubyXL
* Usage: xlsx2yaml [options] [XLSX]
* -s, --site. Prints site information. Includes also id, version and date
* -e, --extended. Prints extended version information:
  constructor, reviewed, approved, created-date and rsmp-version
* If using the -s flag in combination with the -e flag,
  then also the ntsObjectId field is added
* Since the "values" fields in alarms, statuses and commands cannot easily
  be converted from the SXL in Excel format, the "range" field is added if type
  is not boolean and there is no predefined values to choose from.
  Enable using the -e flag
* Typical usage:
  Output to rsmp_schema: No extra options needed
  Output to rst-format for the SXL TLC specification: Use the -e flag (for "value")
  Output to yaml-format and back to Excel-format: Use the -e and -s flag

Notes about yaml2xlsx
---------------------

* Requires: gem install rubyXL
* Needs an excel template
* Prints extended version information if the following fields are present:
  constructor, reviewed, approved, created-date and rsmp-version
  id, version and date
* Since the "values" fields in alarms, statuses and commands cannot easily
  be converted from the SXL in Excel format, the "range" fields is also
  supported.
* The Excel file is written to "output.xlsx"


Notes about yaml2rst
--------------------

* Requires: pip3 install tabulate --user (or apt install python3-tabulate)
* Prints extended version information if the following fields are present:
  constructor, reviewed, approved, created-date and rsmp-version
  id, version and date
  They are generated with xlsx2yaml.rb using the -e and -s flags
* Since the "values" fields in alarms, statuses and commands cannot easily
  be converted from the SXL in Excel format, the "range" fields is also
  supported. The "range" field can be added using xlsx2yaml using the -e flag

Example usages
--------------

Example 1: Convert the SXL from Excel to YAML

```
xlsx2yaml.rb SXL_Traffic_Controller.xlsx
```

Example 2: Convert the SXL from Excel format to RST, including extended attributes

```
xlsx2yaml.rb -e SXL_Traffic_Controller.xlsx | yaml2rst.py > sxl_traffic_light_controller.rst
```

Example 3: Convert the SXL from Excel format to YAML, and then back again to Excel
           Includes extended attributes and site information

```
xlsx2yaml.rb -s -e SXL_Traffic_Controller.xlsx | yaml2xlsx.rb --template "RSMP_Template_SignalExchangeList-20120117.xlsx"
```

