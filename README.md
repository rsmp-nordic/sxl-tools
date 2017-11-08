sxl-tools
=========

Exports signal exchange lists (SXL) used by RSMP to various formats.

* **sxl2csv.pl** - Exports SXL from XLSX format to CSV files, zip compressed.
                   Suitable for use with the RSMP simulators
* **sxl2md.pl** - Exports SXL in XLSX or CSV format to Markdown

Ideas:
* **ots2-clean.pl** - 1.0.7 SXL, adapt according to cId/siteid and num SG/DET

Requires Spreadsheet::XLSX
* Ubuntu: `sudo apt install libspreadsheet-xlsx-perl`
* Arch Linux: `pacman -S zip` and from AUR: `perl-spreadsheet-xlsx` and
  dependencies
