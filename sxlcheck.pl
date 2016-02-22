#!/usr/bin/perl -w
# sxlcheck.pl [SXL in xlsx-format]
# Outputs basic SXL information
# Requires Spreadsheet::XLSX (sudo apt-get install libspreadsheet-xlsx-perl)

# TODO: Test for spaces before or after
# TODO: Test for incorrect cId matches between single and grouped objects

use strict;
use Spreadsheet::XLSX;
use Getopt::Long;

my $all; # Print all details
GetOptions("all" => \$all);

my @files = @ARGV;
my $fname;
my $sheet;

foreach $fname (@files) {
	read_sxl($fname);
}

sub read_sxl {
	my $fname = shift;
	my $workbook = Spreadsheet::XLSX->new($fname);
	foreach $sheet (@{$workbook->{Worksheet}}) {
		if($sheet->{Name} eq "Version") {
			cprint($sheet, 3, 1, "Plant Id:      ");
			cprint($sheet, 5, 1, "Plant Name:    ");
			cprint($sheet, 9, 1, "Constructor:   ") if defined($all);
			cprint($sheet, 11,1, "Reviewed:      ") if defined($all);
			cprint($sheet, 14,1, "Approved:      ") if defined($all);
			cprint($sheet, 17,1, "Created date:  ");
			cprint($sheet, 20,1, "SXL revision:  ");
			cprint($sheet, 20,2, "Revision date: ");
			cprint($sheet, 25,1, "RSMP version:  ") if defined($all);
		} elsif($sheet->{Name} eq "Object types") {;
		} elsif($sheet->{Name} eq "Aggregated status") {;
		} elsif($sheet->{Name} eq "Alarms") {;
		} elsif($sheet->{Name} eq "Status") {;
		} elsif($sheet->{Name} eq "Commands") {;
		} else {
			# Object sheet
			printf "Sheet: ".$sheet->{Name}."\n" if defined($all);
			cprint($sheet, 1,1, "SiteId:        ");
			cprint($sheet, 1,2, "Description:   ");
			
			# Print all grouped objects
			printf "Grouped objects\n";
			my $y = 6;
			while (test($sheet, $y)) {
				if (defined($all)) {
					oprint($sheet, $y);
				}
				$y++;
			}

			# if not "all". Just print the last grouped object
			unless (defined($all)) {
				oprint($sheet, $y-1);
			}
	
			# Print all single objects
			printf "Single objects\n";
			$y = 24;
			while (test($sheet, $y)) {
				oprint($sheet, $y) if defined($all);
				$y++;
			}

			# if not "all". Just print the last single object
			oprint($sheet, $y-1) unless (defined($all));
		}
	}
}

# Cell print
sub cprint {
	my $sheet = shift;
	my $y = shift;
	my $x = shift;
	my $text = shift;
	my $val = $sheet->{Cells}[$y][$x]->{Val};
	if (defined($val)) {
		printf $text.$val."\n";
	} else {
		printf "$text(not defined)\n";
	}
}

# Object print
sub oprint {
	my $sheet = shift;
	my $y = shift;

	# Object, componentId, NTSObjectId
	my $object = $sheet->{Cells}[$y][1]->{Val};
	my $cId = $sheet->{Cells}[$y][2]->{Val};
	my $ntsoId = $sheet->{Cells}[$y][3]->{Val};
	my $externalNtsId = $sheet->{Cells}[$y][4]->{Val};

	unless (defined($object) and defined($cId) and defined($ntsoId)) {
		print "WARNING: row $y incomplete\n";
	} else {
		printf "$object $cId $ntsoId\n";
	}
}

# Test for contens
sub test {
	my $sheet = shift;
	my $y = shift;
	return defined($sheet->{Cells}[$y][0]->{Val});
}
