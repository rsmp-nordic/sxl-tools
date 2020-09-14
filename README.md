sxl-tools
=========

Export sxl
----------

Exports signal exchange lists (SXL) used by RSMP from Excel-format (xlsx) to
various formats.

* **sxl2csv.pl** - Exports SXL from XLSX format to CSV files, zip compressed.
                   Suitable for use with the RSMP simulators
* **sxl2rst.pl** - Exports SXL in XLSX or CSV format to ReStructuredText

Excel macro:
* **rsmp_clean.xlsx** - 1.0.7 SXL, adapt according to cId/siteid and num SG/DET

Requires Spreadsheet::XLSX
* Ubuntu: `sudo apt install libspreadsheet-xlsx-perl`
* Arch Linux: `pacman -S zip` and from AUR: `perl-spreadsheet-xlsx` and
  dependencies
