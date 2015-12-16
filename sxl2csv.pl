#!/usr/bin/perl -w
# sxl2csv.pl [SXL in xlsx-format]
# outputs csv-file for each worksheet and saves to zip file 'Objects.zip'
# Requires Spreadsheet::XLSX (sudo apt-get install libspreadsheet-xlsx-perl)

#NOTE: Supports only XLSX-format as input. See Spreadsheet::ParseExcel
#NOTE: Untested i18n support. Encoding
#NOTE: Add option to adjust the SXL and save

use strict;
use Spreadsheet::XLSX;

my $fname = shift;
my $oname;
my $value;
my $workbook = Spreadsheet::XLSX->new($fname);

system("mkdir Objects");
foreach my $sheet (@{$workbook->{Worksheet}}) {
    $oname = $sheet->{Name}.".csv";
    # Add underscore instead of whitespace in filenames
    $oname =~ tr/ /_/;
    $sheet -> {MaxRow} ||= $sheet -> {MinRow};

    open(FILE, ">Objects/$oname") or die "Cannot open file";

    foreach my $row (0 .. $sheet->{MaxRow}) { 
        $sheet->{MaxCol} ||= $sheet->{MinCol};
        foreach my $col (0 ..  $sheet->{MaxCol}) {
            my $cell = $sheet->{Cells}[$row][$col];

            if ($cell) {
                $value = $cell->{Val};
                if ($value =~ /\r/) {
                    # Remove any carriage returns that might exist in cells
                    $value =~ tr/\r//d;
                    # Excel quotes the value if it contains newlines
                    $value = "\"$value\"";
                }
                printf FILE "%s;", $value;
            } else {
                printf FILE ";";
            }
        }
        printf FILE "\r\n";
    }
}
system("zip -q -j Objects.zip Objects/*");
system("rm -rf Objects");
