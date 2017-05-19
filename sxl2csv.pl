#!/usr/bin/perl -w
# sxl2csv.pl [SXL in xlsx-format]
# outputs csv-file for each worksheet and saves to zip file 'Objects.zip'
# Requires Spreadsheet::XLSX and zip
#	Ubuntu: $ sudo apt-get install libspreadsheet-xlsx-perl)
#	Arch: $ https://aur.archlinux.org/perl-spreadsheet-xlsx.git and dependencies...
#		  $ pacman -S zip

#NOTE: Supports only XLSX-format as input. See Spreadsheet::ParseExcel
#NOTE: Untested i18n support. Encoding
#NOTE: Add option to modify the SXL and save
#TODO: Add option to rename zip file after source file

use strict;
use Spreadsheet::XLSX;

my $fname = shift;
my $oname;
my $value;
my $semi;
my $workbook = Spreadsheet::XLSX->new($fname);

system("mkdir Objects");
foreach my $sheet (@{$workbook->{Worksheet}}) {
	$oname = $sheet->{Name}.".csv";
	# Add underscore instead of whitespace in filenames
	$oname =~ tr/ /_/;
	$sheet -> {MaxRow} ||= $sheet -> {MinRow};

	open(FILE, ">Objects/$oname") or die "Cannot open file";

	foreach my $row (0 .. $sheet->{MaxRow}) { 
	$semi=0;
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
				$semi++;
			} else {
				printf FILE ";";
				$semi++;
			}
		}

		# FIXME: Workaround for a special case
		# For some implementations, at least seven semi colon are required
		if ($sheet->{Name} =~ /Version/) {
			while($semi < 7) {
				printf FILE ";";
				$semi++;
			}
		}
		printf FILE "\r\n";
	}

	# FIXME: Workaround for a special case
	# For some implementations, two extra empty rows are required
	if ($sheet->{Name} =~ /Version/) {
		printf FILE ";;;;;;;\r\n";
		printf FILE ";;;;;;;\r\n";
	}
}
system("zip -q -j Objects.zip Objects/*");
system("rm -rf Objects");

# FIXME: Rename after source name
$fname =~ s/.xlsx//;
system("mv Objects.zip $fname.zip");

