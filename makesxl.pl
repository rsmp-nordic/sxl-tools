#!perl
# Exports signal exchange lists (SXL) in Excel-format to CSV files
# Requires ActiveState Perl, Win32::OLE and MS Excel

# Copyright (C) 2012,2013 David Otterdahl <david.otterdahl@gmail.com>

use Win32::OLE qw(EVENTS);
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Win32::OLE::NLS qw(:LOCALE :DATE);
use Getopt::Long;

$Win32::OLE::Warn = 3; # Die on Errors.

my $obj_path;
my $excelfile;
my $visible;
GetOptions(
	"output=s" => \$obj_path,
	"input=s" => \$excelfile,
	"visible" => \$visible,
	);

unless (defined($obj_path) and defined($excelfile)) {
	print "usage: makesxl.sl --input [sxl-file] --output [path] [--visible]";
	exit 1;
}

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
	|| Win32::OLE->new('Excel.Application', 'Quit');

$Excel->{DisplayAlerts}=0;

my $Book = $Excel->Workbooks->Open($excelfile);

# Make Excel visible during export (disabled by default)
if (defined($visible)) {
	$Excel->{Visible} = 1;
}

system "mkdir \"$obj_path\\Objects\"";

my $sheetcnt = $Book->Worksheets->Count();
foreach (1..$sheetcnt) {
	my $name = $Book->Worksheets($_)->{Name};
	$name =~ s/\ /_/g;
	$name = "$obj_path\\Objects\\$name.csv";
	$Book->Worksheets($_)->Activate();
	$Book->SaveAs($name, xlCSV);
}

undef $Book;
undef $Excel;
