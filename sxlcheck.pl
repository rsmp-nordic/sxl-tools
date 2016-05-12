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
			# Sanity check. Check for semicolon in the "Comment" field for the first return value
			# TODO: Check in all the return values; not just the first one
			my $y = 6;
			while (test($sheet, $y, 7)) {
				semi_check($sheet, 7, $y);
				$y++;
			}
		} elsif($sheet->{Name} eq "Commands") {;

			# Sanity check. Check for semicolon in the "Comment" field for the first
			# argument in all command sections
			# TODO: Check in all the arguments values; not just the first one
			my $sec;
			my $y;
			my @sections = command_section($sheet);
			foreach $sec (@sections) {
				# Need to check each command section
				$y = $sec;
				while (test($sheet, $y, 8)) {
					semi_check($sheet, 8, $y);
					$y++;
				}
			}
		} else {
			# Object sheet
			printf "Sheet: ".$sheet->{Name}."\n" if defined($all);
			cprint($sheet, 1,1, "SiteId:        ");
			cprint($sheet, 1,2, "Description:   ");

			# Print all grouped objects
			printf "Grouped objects\n";
			my $y = 6;
			while (test($sheet, $y, 0)) {
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
			while (test($sheet, $y, 0)) {
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

# Find command section
# Return a list of y-positions for the start of each section
sub command_section {
	my $sheet = shift;
	my $y = 4; # Section won't start before row 4
	my @list;
	my $text;
	while ($y<100) {
		$text = "";
		$text = $sheet->{Cells}[$y][0]->{Val} if(test($sheet, $y, 0));

		# We're adding +2 because that's where the actual command starts
		if($text =~ /Functional position/) {
			push @list, $y+2;
		} elsif( $text =~ /Functional state/) {
			push @list, $y+2;
		} elsif($text =~ /Manouver/) {
			push @list, $y+2;
		} elsif($text =~ /Parameter/) {
			push @list, $y+2;
			return @list;
		} else {
		}
		$y++;
	}
	print "Error: did not find all command sections\n";
	return;
}

# Semicolon check
sub semi_check {
	my $sheet = shift;
	my $x = shift;
	my $y = shift;
	my $comment = $sheet->{Cells}[$y][$x]->{Val};

	printf "WARNING: Found semicolon in comment field: $comment\n" if ($comment =~ /;/);
}

# Test for contens if the first column
sub test {
	my $sheet = shift;
	my $y = shift;
	my $x = shift;
	return defined($sheet->{Cells}[$y][$x]->{Val});
}
