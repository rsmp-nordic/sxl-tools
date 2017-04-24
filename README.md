sxl-tools
=========

Exports signal exchange lists (SXL) used by RSMP from Excel-format (xlsx) to
various formats.

Requires Spreadsheet::XLSX
* Ubuntu: `sudo apt install libspreadsheet-xlsx-perl`
* Arch Linux: `pacman -S zip` and `perl-spreadsheet-xlsx` and dependencies
  from AUR

* sxl2csv.pl - Exports to CSV files. Suitable for use with the RSMP simulators
* sxl2md.pl - Exports to Markdown
