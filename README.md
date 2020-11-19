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
* **rsmp_clean.xlsm** - 1.0.7 SXL, adapt according to cId/siteid and num SG/DET

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
* -s, --site. Prints [site information](#site). Includes also id, version and date
* -e, --extended. Prints [extended information](#extended)
* If using the -s flag in combination with the -e flag,
  then also the ntsObjectId field is added
* Since the "values" fields in alarms, statuses and commands cannot easily
  be converted from the SXL in Excel format, the "range" field is added if type
  is not boolean and there is no predefined values to choose from.
  Enable using the -e flag
* Typical usage:
  * Output to rsmp_schema: No extra options needed
  * Output to rst-format for the SXL TLC specification: Use the -e flag (for "value")
  * Output to yaml-format and back to Excel-format: Use the -e and -s flag

Notes about yaml2xlsx
---------------------

* Requires: gem install rubyXL
* Needs an excel template
* Prints [extended information](#extended) if the fields are present
* Since the "values" fields in alarms, statuses and commands cannot easily
  be converted from the SXL in Excel format, the "range" fields is also
  supported.
* The Excel file is written to "output.xlsx"


Notes about yaml2rst
--------------------

* Requires: pip3 install tabulate --user (or apt install python3-tabulate)
* Prints some of the [extended information](#extended) if they are present
  They are generated with xlsx2yaml.rb using the -e and -s flags
* Since the "values" fields in alarms, statuses and commands cannot easily
  be converted from the SXL in Excel format, the "range" fields is also
  supported. The "range" field can be added using xlsx2yaml using the -e flag

<a name="site"></a>
Site information
----------------
Site fields contains all the individual components of a specific site. This
is needed by the RSMP simulator to tie a specific component (object) to an
alarm, status or command, but is not needed to construct a JSon schema.

<a name="extended"></a>
Extended information
--------------------
Extended fields contains all fields of the SXL, not just those needed to
construct a JSon schema.

List of fields:
* `constructor` (version)
* `reviewed` (version)
* `approved` (version)
* `created-date` (version)
* `rsmp-version` (version)
* `ntsObjectId` (objects) only if site information is enabled
* `externalNtsId` (objects) only if site information is enabled
* `range` (alarms, status, commands)
* `object` (alarms, status, commands) only if site information is enabled
* `externalAlarmCodeId` (alarms)
* `externalNtsAlarmCodeId` (alarms)
* `functional_position` (aggregated status)
* `functional_state` (aggregated status)

Mapping between XLSX and YAML format of the SXL
-----------------------------------------------
Version sheet

| Cell 	| Name in Excel   	| YAML         	| extended? 	|
|------	|-----------------	|--------------	|-----------	|
| B2   	| Plant id        	| id           	|           	|
| B6   	| Plant name      	| description  	|           	|
| B10  	| Constructor     	| constructor  	| yes       	|
| B12  	| Reviewed        	| reviewed     	| yes       	|
| B15  	| Approved        	| approved     	| yes       	|
| B18  	| Created date    	| created-date 	| yes       	|
| B21  	| Revision number 	| version      	|           	|
| C21  	| Revision date   	| date         	|           	|
| B26  	| RSMP version    	| rsmp-version 	| yes       	|

Object types

| Cell 	| Name in Excel       	| YAML         	| extended? 	|
|------	|---------------------	|--------------	|-----------	|
| A7.. 	| ObjectType          	| [ObjectType] 	|           	|
| B7.. 	| Description/comment 	| description  	| yes       	|

Objects ([site information](#site))

| Cell 	| Name in Excel 	| YAML                    	| extended? 	|
|------	|---------------	|-------------------------	|-----------	|
| A7.. 	| ObjectType    	| [ObjectType]            	|           	|
| B7.. 	| Object        	| [Object]                	|           	|
| C7.. 	| componentId   	| [Object]: [componentId] 	|           	|
| D7.. 	| NTSObjectId   	| ntsObjectid             	| yes       	|
| E7.. 	| externalNtsId 	| externalNtsId           	| yes       	|
| F7.. 	| Description   	| description             	| yes       	|

Aggregated status

| Cell  	| Name in Excel      	| YAML                	| extended? 	|
|-------	|--------------------	|---------------------	|-----------	|
| C7..  	| functionalPosition 	| functional_position 	| yes       	|
| D7..  	| functionalState    	| functionalState     	| yes       	|
| A17.. 	| Comment            	| description         	| yes       	|

Alarms

| Cell  	| Name in Excel          	| YAML                   	| extended? 	|
|-------	|------------------------	|------------------------	|-----------	|
| A7..  	| ObjectType             	| [ObjectType]           	|           	|
| B7..  	| Object (optional)      	| [Object]               	| yes       	|
| C7..  	| alarmCodeId            	| [alarmCodeId]          	|           	|
| D7..  	| Description            	| description            	|           	|
| E7..  	| externalAlarmCodeId    	| externalAlarmCodeId    	| yes       	|
| F7..  	| externalNtsAlarmCodeId 	| externalNtsAlarmCodeId 	| yes       	|
| G7..  	| Priority               	| priority               	|           	|
| H7..  	| Category               	| category               	|           	|
| I7..  	| Name                   	| [Name]                 	|           	|
| J7..  	| Type                   	| type                   	|           	|
| K7..  	| Value                  	| values (list)          	|           	|
| K7..  	| Value                  	| range (if not boolean) 	| yes       	|
| L7..  	| Comment                	| description            	|           	|

Status

| Cell  	| Name in Excel     	| YAML                   	| extended? 	|
|-------	|-------------------	|------------------------	|-----------	|
| A7..  	| ObjectType        	| [ObjectType]           	|           	|
| B7..  	| Object (optional) 	| [Object]               	| yes       	|
| C7..  	| statusCodeId      	| [statusCodeId]         	|           	|
| D7..  	| Description       	| description            	|           	|
| E7..  	| Name              	| [Name]                 	|           	|
| F7..  	| Type              	| type                   	|           	|
| G7..  	| Value             	| values (list)          	|           	|
| G7..  	| Value             	| range (if not boolean) 	| yes       	|
| H7..  	| Comment           	| description            	|           	|

Commands

| Cell  	| Name in Excel     	| YAML                   	| extended? 	|
|-------	|-------------------	|------------------------	|-----------	|
| A7..  	| ObjectType        	| [ObjectType]           	|           	|
| B7..  	| Object (optional) 	| [Object]               	| yes       	|
| C7..  	| commandCodeId     	| [commandCodeId]        	|           	|
| D7..  	| Description       	| description            	|           	|
| E7..  	| Name              	| [Name]                 	|           	|
| F7..  	| Command           	| command                	|           	|
| G7..  	| Type              	| type                   	|           	|
| H7..  	| Value             	| values (list)          	|           	|
| H7..  	| Value             	| range (if not boolean) 	| yes       	|
| I7..  	| Comment           	| description            	|           	|

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

