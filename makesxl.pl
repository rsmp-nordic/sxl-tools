#!/usr/bin/perl -w
# Exports signal exchange lists (SXL) in Excel-format to CSV files
# Requires MS Excel and Windows

# Copyright (C) 2012 David Otterdahl <david.otterdahl@gmail.com>
#
# This program is free software; you can redistribute it and/or modify it
# under the terms of the GNU General Public License as published by the
# Free Software Foundation version 2 of the License.
#
# This program is distributed in the hope that it will be useful, but
# WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
# General Public License for more details.
#
# You should have received a copy of the GNU General Public License along
# with this program; if not, write to the Free Software Foundation, Inc.,
# 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301, USA.

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
