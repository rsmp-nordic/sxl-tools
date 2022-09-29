sxl-tools
=========

Exports signal exchange lists (SXL) used by RSMP from Excel-format (xlsx) to
various formats.

* **create_template.py** - Creates SXL template in Excel format
* **xlsx2csv.rb**  - Reads SXL in Excel format and outputs to CSV format
* **xlsx2yaml.rb** - Reads SXL in Excel format and outputs to YAML format
* **yaml2xlsx.rb** - Reads SXL in YAML format and outputs to Excel format
* **yaml2rst.py**  - Reads SXL in YAML format and outputs to RST format

Notes about create_template.py
------------------------------
* Requires: pip3 install xlsxwriter --user
* Usage: create_template.py [OPTIONS]
* See create_template.py -h for available options

Notes about xlsx2csv
--------------------

* Requires: gem install rubyXL
* Usage: xlsx2csv [XLSX]
* Exports SXL from XLSX format to CSV files, zip compressed.
  Suitable for use with the RSMP simulators

Notes about xlsx2yaml
---------------------

* Requires: gem install rubyXL
* Usage: xlsx2yaml [options] [XLSX]
* -s, --site. Prints [site information](#site). Includes also id, version and date
* Typical usage:
  * Output to rsmp_schema: No extra options needed
  * Output to rst-format for the SXL TLC specification: No extra options needed
  * Output to yaml-format and back to Excel-format: Use -s flag

Notes about yaml2xlsx
---------------------

* Requires: gem install rubyXL
* Needs an excel template
* The Excel file is written to "output.xlsx"


Notes about yaml2rst
--------------------

* Requires: pip3 install pyyaml tabulate --user (or apt install python3-tabulate)
* Prints extended fields if the "--extended" option is used.

<a name="site"></a>
Site information
----------------
Site fields contains all the individual components of a specific site. This
is needed by the RSMP simulator to tie a specific component (object) to an
alarm, status or command, but is not needed to construct a JSon schema.

Mapping between XLSX and YAML format of the SXL
-----------------------------------------------
Version sheet

| Cell 	| Name in Excel   	| YAML         	|
|------	|-----------------	|--------------	|
| B2   	| Plant id        	| id           	|
| B6   	| Plant name      	| description  	|
| B10  	| Constructor     	| constructor  	|
| B12  	| Reviewed        	| reviewed     	|
| B15  	| Approved        	| approved     	|
| B18  	| Created date    	| created-date 	|
| B21  	| Revision number 	| version      	|
| C21  	| Revision date   	| date         	|
| B26  	| RSMP version    	| rsmp-version 	|

Object types

| Cell 	| Name in Excel       	| YAML         	|
|------	|---------------------	|--------------	|
| A7.. 	| ObjectType          	| [ObjectType] 	|
| B7.. 	| Description/comment 	| description  	|

Objects ([site information](#site))

| Cell 	| Name in Excel 	| YAML                    	|
|------	|---------------	|-------------------------	|
| A7.. 	| ObjectType    	| [ObjectType]            	|
| B7.. 	| Object        	| [Object]                	|
| C7.. 	| componentId   	| [Object]: [componentId] 	|
| D7.. 	| NTSObjectId   	| ntsObjectid             	|
| E7.. 	| externalNtsId 	| externalNtsId           	|
| F7.. 	| Description   	| description             	|

Aggregated status

| Cell  	| Name in Excel      	| YAML                          |
|-------	|--------------------	|-----------------------------  |
| C7..  	| functionalPosition 	| functional_position           |
| D7..  	| functionalState    	| functional_state              |
| A17.. 	| Comment            	| aggregated_status_description |

Alarms

| Cell  	| Name in Excel          	| YAML                   		|
|-------	|------------------------	|------------------------		|
| A7..  	| ObjectType             	| [ObjectType]           		|
| B7..  	| Object (optional)      	| [Object]               		|
| C7..  	| alarmCodeId            	| [alarmCodeId]          		|
| D7..  	| Description            	| description            		|
| E7..  	| externalAlarmCodeId    	| externalAlarmCodeId    		|
| F7..  	| externalNtsAlarmCodeId 	| externalNtsAlarmCodeId 		|
| G7..  	| Priority               	| priority               		|
| H7..  	| Category               	| category               		|
| I7..  	| Name                   	| [Name]				|
| J7..  	| Type                   	| type					|
| K7..  	| Value                  	| values (list)          		|
| K7..  	| Value                  	| max,min (integer, long, real)		|
| L7..  	| Comment                	| description            		|

Status

| Cell  	| Name in Excel     	| YAML                   		|
|-------	|-------------------	|------------------------		|
| A7..  	| ObjectType        	| [ObjectType]           		|
| B7..  	| Object (optional) 	| [Object]               		|
| C7..  	| statusCodeId      	| [statusCodeId]         		|
| D7..  	| Description       	| description            		|
| E7..  	| Name              	| [Name]                 		|
| F7..  	| Type              	| type                   		|
| G7..  	| Value             	| values (list)          		|
| G7..  	| Value             	| max,min (integer, long, real)		|
| H7..  	| Comment           	| description            		|

Commands

| Cell  	| Name in Excel     	| YAML                   		|
|-------	|-------------------	|------------------------		|
| A7..  	| ObjectType        	| [ObjectType]           		|
| B7..  	| Object (optional) 	| [Object]               		|
| C7..  	| commandCodeId     	| [commandCodeId]        		|
| D7..  	| Description       	| description            		|
| E7..  	| Name              	| [Name]                 		|
| F7..  	| Command           	| command                		|
| G7..  	| Type              	| type                   		|
| H7..  	| Value             	| values (list)          		|
| h7..  	| Value             	| max,min (integer, long, real) 	|
| I7..  	| Comment           	| description            		|

Example usages
--------------

Example 1: Convert the SXL from Excel to YAML.

```
xlsx2yaml.rb SXL_Traffic_Controller.xlsx
```

Example 2: Convert the SXL from Excel format to RST.

```
xlsx2yaml.rb SXL_Traffic_Controller.xlsx | yaml2rst.py > sxl_traffic_light_controller.rst
```

Example 3: Convert the SXL from Excel format to YAML, and then back again to Excel using a template.
Includes site information

```
xlsx2yaml.rb -s SXL_Traffic_Controller.xlsx | yaml2xlsx.rb --template "RSMP_Template_SignalExchangeList-20120117.xlsx"
```

