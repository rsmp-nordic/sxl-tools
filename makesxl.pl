#!perl
#!/usr/bin/perl -w
# Exports signal exchange lists (SXL) in Excel-format to CSV files
# Requires ActiveState Perl, Win32::OLE and MS Excel
# TODO: Requires "Objects" directory
# TODO: Doesn't quit Excel

# Copyright (C) 2012,2013 David Otterdahl <david.otterdahl@gmail.com>

use Win32::OLE qw(EVENTS);
use Win32::OLE::Const 'Microsoft Excel';
use Win32::OLE::Variant;
use Win32::OLE::NLS qw(:LOCALE :DATE);

$Win32::OLE::Warn = 3; # Die on Errors.

# Change $obj_path to something useful e.g.:
# 'C:\\Documents and Settings\\User\\Desktop\\'
my $obj_path = 'C:\\tmp\\';
my $excelfile = 'SXL.xlsx';

my $Excel = Win32::OLE->GetActiveObject('Excel.Application')
	|| Win32::OLE->new('Excel.Application', 'Quit');

$Excel->{DisplayAlerts}=0;

my $Book = $Excel->Workbooks->Open($obj_path.$excelfile);

# Make Excel visible during export (disabled by default)
#$Excel->{Visible} = 1;

my $sheetcnt = $Book->Worksheets->Count();
foreach (1..$sheetcnt) {
	my $name = $Book->Worksheets($_)->{Name};
	$name =~ s/\ /_/g;
	$Book->Worksheets($_)->Activate();
	$Book->SaveAs($obj_path.'Objects\\'.$name.'.csv', xlCSV);
}

undef $Book;
undef $Excel;
