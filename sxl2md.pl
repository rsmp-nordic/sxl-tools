#!/usr/bin/perl -w
# sxlcheck.pl [SXL in xlsx-format]
# Outputs basic SXL information
# Requires Spreadsheet::XLSX (sudo apt-get install libspreadsheet-xlsx-perl)

# TODO: Test for spaces before or after
# TODO: Test for incorrect cId matches between single and grouped objects
# TODO: Add dates

use strict;
use Spreadsheet::XLSX;
use Getopt::Long;

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
			print_version($sheet);
		} elsif($sheet->{Name} eq "Object types") {
			print_object_types($sheet);
		} elsif($sheet->{Name} eq "Aggregated status") {
			print_aggregated_status($sheet);
		} elsif($sheet->{Name} eq "Alarms") {
			print_alarms($sheet);
		} elsif($sheet->{Name} eq "Status") {
			print_status($sheet);
		} elsif($sheet->{Name} eq "Commands") {
			print_commands($sheet);
		} else { # Objects
			print_objects($sheet);
		}
	}
}

sub print_version {
	my $sheet = shift;
	printf("# Signal Exchange List\n");

	cprint($sheet, 3, 1, "Plant Id");
	cprint($sheet, 5, 1, "Plant Name");
	cprint($sheet, 9, 1, "Constructor");
	cprint($sheet, 11,1, "Reviewed");
	cprint($sheet, 14,1, "Approved");
	cprint($sheet, 17,1, "Created date");
	cprint($sheet, 20,1, "SXL revision");
	cprint($sheet, 20,2, "Revision date");
	cprint($sheet, 25,1, "RSMP version");
	printf("\nSections\n");
	printf("--------\n");
	printf("+ [Object types](#object_types)\n");
	printf("+ [Objects](#objects)\n");
	printf("+ [Aggregated status](#aggregated_status)\n");
	printf("+ [Alarms](#alarms)\n");
	printf("+ [Status](#status)\n");
	printf("+ [Commands](#commands)\n");
}

sub print_object_types {
	my $sheet = shift;
	# Object types sheet
	printf("<a id=\"object_types\"></a>\n");
	printf("\nObject Types\n");
	printf("============\n");

	# Print all grouped objects
	printf "\nGrouped objects\n";
	printf "---------------\n";
	printf "|ObjectType|Description|\n";
	printf "|----------|-----------|\n";
	my $y = 6;
	while (test($sheet, $y, 0)) {
		otprint($sheet, $y);
		$y++;
	}

	# Print all single objects
	printf "\nSingle objects\n";
	printf "--------------\n";
	printf "|ObjectType|Description|\n";
	printf "|----------|-----------|\n";
	$y = 18;
	while (test($sheet, $y, 0)) {
		otprint($sheet, $y);
		$y++;
	}
}

sub print_aggregated_status {
	my $sheet = shift;
	# Aggregated status sheet
	printf("<a id=\"aggregated_status\"></a>\n");
	printf("\nAggregated status per grouped object\n");
	printf("====================================\n");

	# Print all grouped objects
	printf "|ObjectType|Status|functionalPosition|functionalState|Description|\n";
	printf "|----------|------|------------------|---------------|-----------|\n";
	my $y = 6;
	while (test($sheet, $y, 0)) {
		aggprint($sheet, $y);
		$y++;
	}

	# Print all state bits
	printf "\n|State- Bit nr (12345678)|Description|Comment|\n";
	printf "|------------------------|-----------|-------|\n";
	stateprint($sheet);

}

sub print_objects {
	my $sheet = shift;
	# Object sheet
	printf("<a id=\"objects\"></a>\n");
	printf("\nSite Objects\n");
	printf("============\n");
	cprint($sheet, 1,1, "SiteId");
	cprint($sheet, 1,2, "Description");

	# Print all grouped objects
	printf "\nGrouped objects\n";
	printf "---------------\n";
	printf "|ObjectType|Object|componentId|NTSObjectId|externalNtsId|Description|\n";
	printf "|----------|------|-----------|-----------|-------------|-----------|\n";
	my $y = 6;
	while (test($sheet, $y, 0)) {
		oprint($sheet, $y);
		$y++;
	}

	# Print all single objects
	printf "\nSingle objects\n";
	printf "--------------\n";
	printf "|ObjectType|Object|componentId|NTSObjectId|externalNtsId|Description|\n";
	printf "|----------|------|-----------|-----------|-------------|-----------|\n";
	$y = 24;
	while (test($sheet, $y, 0)) {
		oprint($sheet, $y);
		$y++;
	}
}

sub print_alarms {
	my $sheet = shift;
	printf("<a id=\"alarms\"></a>\n");
	printf("\n# Alarms\n");
	
	# Print header
	my $i;
	printf "| ObjectType | Object (optional) | alarmCodeId | Description | externalAlarmCodeId | externalNtsAlarmCodeId | Priority | Category |\n";
	printf "| ---------- | ----------------- |:-----------:| ----------- | ------------------- | ---------------------- |:--------:|:--------:|\n";

	# Print alarms
	my $y = 6;
	my $return_text = "";
	while (test($sheet, $y, 7)) {
		my $has_return_values = get_no_return_values($sheet, $y, 8, 4);
		aprint($sheet, $y, 8, $has_return_values);
		print "\n";

		# Collect return values
		if($has_return_values > 0) {
			my $xCodeId = $sheet->{Cells}[$y][2]->{Val};

			# Print header
			$return_text .= "\n<a id=\"$xCodeId\"></a>";
			$return_text .= "\n## Return Values for $xCodeId\n";
			$return_text .= "|Name|Type|Value|Comment|\n";
			$return_text .= "|----|----|-----|-------|\n";

			$return_text = rprint($sheet, $y, 8, 4, 2, $return_text);
		}
		$y++;
	}

	# Print return values
	print $return_text;
}

sub print_status {
	my $sheet = shift;
	printf("<a id=\"status\"></a>\n");
	printf("\n# Status\n");


	# Print header
	my $i;
	printf "| ObjectType | Object (optional) | statusCodeId | Description |\n";
	printf "| ---------- | ----------------- |:------------:| ----------- |\n";
	
	# Print status
	my $y = 6;
	my $return_text = "";
	while (test($sheet, $y, 7)) {
		my $has_return_values = get_no_return_values($sheet, $y, 4, 4);
		aprint($sheet, $y, 4, $has_return_values);
		print "\n";

		# Collect return values
		if($has_return_values > 0) {
			my $xCodeId = $sheet->{Cells}[$y][2]->{Val};

			# Print header
			$return_text .= "\n<a id=\"$xCodeId\"></a>";
			$return_text .= "\n## Return Values for $xCodeId\n";
			$return_text .= "|Name|Type|Value|Comment|\n";
			$return_text .= "|----|----|-----|-------|\n";

			$return_text = rprint($sheet, $y, 4, 4, 2, $return_text);
		}
		$y++;
	}

	# Print return values
	print $return_text;
}

sub print_commands {
	my $sheet = shift;
	printf("<a id=\"commands\"></a>\n");
	printf("\n# Commands\n");

	my $sec;
	my $y;
	my @sections = command_section($sheet);

	# Print header
	my $i;
	printf "| ObjectType | Object (optional) | commandCodeId | Description |\n";
	printf "| ---------- | ----------------- |:-------------:| ----------- |\n";

	my $txt = "";
	my $return_text = "";
	foreach $sec (@sections) {
		# Need to check each command section
		$y = $sec;

		# Print command
		while (test($sheet, $y, 7)) {
			my $has_return_values = get_no_return_values($sheet, $y, 4, 5);
			aprint($sheet, $y, 4, $has_return_values);
			print "\n";

			# Collect return values
			if($has_return_values > 0) {
				my $xCodeId = $sheet->{Cells}[$y][2]->{Val};

				# Print header
				$txt = "\n<a id=\"$xCodeId\"></a>";
				$txt .= "\n## Arguments for $xCodeId\n";
				$txt .= "|Name|Command|Type|Value|Comment|\n";
				$txt .= "|----|-------|----|-----|-------|\n";

				$return_text .= rprint($sheet, $y, 4, 5, 3, $txt);

			}
			$y++;
		}
	}

	# Print return values
	print $return_text;
	print "\n";
}


# Cell print
sub cprint {
	my $sheet = shift;
	my $y = shift;
	my $x = shift;
	my $text = shift;
	my $val = $sheet->{Cells}[$y][$x]->{Val};

	printf "+ **$text**:";
        $val =~ s/%/%%/g if(defined($val)); # Needed for printf()
	printf " $val" if(defined($val));
	printf "\n";
}

# Print object type
sub otprint {
	my $sheet = shift;
	my $y = shift;

	my $objecttype = $sheet->{Cells}[$y][0]->{Val};
	my $description = $sheet->{Cells}[$y][5]->{Val};

	printf "|$objecttype|";
	printf "$description|" if(defined($description));
	printf "|\n";
}


# Print object
sub oprint {
	my $sheet = shift;
	my $y = shift;

	my $objecttype = $sheet->{Cells}[$y][0]->{Val};
	my $object = $sheet->{Cells}[$y][1]->{Val};
	my $cId = $sheet->{Cells}[$y][2]->{Val};
	my $ntsoId = $sheet->{Cells}[$y][3]->{Val};
	my $externalNtsId = $sheet->{Cells}[$y][4]->{Val};
	my $description = $sheet->{Cells}[$y][5]->{Val};

	unless (defined($object) and defined($cId) and defined($ntsoId)) {
		print STDERR "WARNING: row $y incomplete\n";
	} else {
		printf "|$objecttype|$object|$cId|$ntsoId|";

		# ExternalNtsId
		printf "$externalNtsId" if(defined($externalNtsId));
		printf "|";

		printf "$description" if(defined($description));
		printf "|\n";
	}
}

# Print aggregated status
sub aggprint {
	my $sheet = shift;
	my $y = shift;

	my $objecttype = $sheet->{Cells}[$y][0]->{Val};
	my $state = $sheet->{Cells}[$y][1]->{Val};
	my $functionalPosition = $sheet->{Cells}[$y][2]->{Val};
	my $functionalState = $sheet->{Cells}[$y][3]->{Val};
	my $description = $sheet->{Cells}[$y][4]->{Val};

	printf "|$objecttype|";

	printf "$state" if(defined($state));
	printf "|";

	printf "$functionalPosition" if(defined($functionalPosition));
	printf "|";

	printf "$functionalState" if(defined($functionalState));
	printf "|";

	printf "$description" if(defined($description));
	printf "|\n";


}

# Print state bits (aggregated status)
sub stateprint {
	my $sheet = shift;
	my $y;
	my @bit;
	for($y = 0; $y < 8; $y++) {
		$bit[$y] = $sheet->{Cells}[$y+16][2]->{Val};

		# Remove line breaks
		if(defined($bit[$y])) {
			$bit[$y] =~ s/\r//g;
			$bit[$y] =~ s/\n/<br>/g;
		}
	}
	printf "|1|Local mode|";          print $bit[0] if(defined($bit[0])); print "|\n";
	printf "|2|No communications|";   print $bit[1] if(defined($bit[1])); print "|\n";
	printf "|3|High priority fault|"; print $bit[2] if(defined($bit[2])); print "|\n";
	printf "|4|Medium priority fault|"; print $bit[3] if(defined($bit[3])); print "|\n";
	printf "|5|Low priority fault|";  print $bit[4] if(defined($bit[4])); print "|\n";
	printf "|6|Connected / Normal - In Use|";  print $bit[5] if(defined($bit[5])); print "|\n";
	printf "|7|Connected / Normal - Idle|";  print $bit[6] if(defined($bit[6])); print "|\n";
	printf "|8|Not Connected|";  print $bit[7] if(defined($bit[7])); print "|\n";

}

# Print alarm/status/commands
sub aprint {
	my $sheet = shift;
	my $y = shift;  # Start row
	my $col_length = shift; # 8 for alarms, 4 for status and commands
	my $has_return_values = shift; # num of return values
	my $col_link = 2; # Which column to make a link, if return values exist

	# Get values for a row
	my $i;
	my $x = 0;
	my @val;
	for($i = 0; $i < $col_length; $i++) {
		$val[$i] = $sheet->{Cells}[$y][$x++]->{Val};

		unless(defined($val[$i])) {
			$val[$i] = "";
		}
	}
	
	# Make column into a link if return values exist
	if($has_return_values > 0) {
		$val[$col_link] = "[".$val[$col_link]."](#".$val[$col_link].")"; 
	}

	# Print row
	print "|";
	for($i = 0; $i < $col_length; $i++) {
		$val[$i] =~ s/\r//g;
		$val[$i] =~ s/\n/<br>/g;
		print "$val[$i]|";
	}
}

# Print arguments/return values
sub rprint {
	my $sheet = shift;
	my $y = shift;  # Start row
	my $start_x = shift; # 8 for alarms, 4 for status and commands
	my $return_value_col_length = shift; # 4 for alarm and status, 5 for commands
	my $value_list_col = shift; # this column of return values/arguments should be split into bullet list, 2 for alarm and status, 3 for commands
	my $return_text = shift;

	my $i;
	my @val;
	my $x = $start_x;
	my $col_length;

	# return values
	while(test($sheet, $y, $x)) {
		# Get values for a row
		$col_length = $return_value_col_length;
		for($i = 0; $i < $col_length; $i++) {
			$val[$i] = $sheet->{Cells}[$y][$x++]->{Val};
			unless(defined($val[$i])) {
				$val[$i] = "";
			}
		}

		# Check for semicolon in the "Comment" field, which is the last one
		semi_check($sheet, $x, $y);

		# Print row
		$return_text .= "|";
		for($i = 0; $i < $col_length; $i++) {
			# 'Value' should be split into bullet list.
			# Markdown don't support bullet lists in a table
			# so use inline HTML instead
			$val[$i] =~ s/\n-/\n/g;	# Remove '-', after line break
			$val[$i] =~ s/^-//g;	# Remove leading '-'
			if($i == $value_list_col) {
				# Find line breaks and convert to them to bullet list in HTML
				if($val[$i] =~ /\r\n/) {
					$val[$i] =~ s/\r//g;
					my @list = split("\n", $val[$i]);
					my $v;
					$val[$i] = "<ul>";
					foreach $v (@list) {
						$val[$i] = $val[$i]."<li>$v</li>";
					}
					$val[$i] = $val[$i]."</ul>";
				}
			}
			else {
				# Remove line breaks
				$val[$i] =~ s/\r//g;
				$val[$i] =~ s/\n/<br>/g;
			}
			$return_text .= "$val[$i]|";
		}
		$return_text .= "\n";
	}
	return $return_text;
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
	if(defined($comment)) {
		printf STDERR "WARNING: Found semicolon in comment field: $comment\n" if ($comment =~ /;/);
	}
}

# Test for contens if the first column
sub test {
	my $sheet = shift;
	my $y = shift;
	my $x = shift;
	return defined($sheet->{Cells}[$y][$x]->{Val});
}

# Get number of arguments/return values for given row
sub get_no_return_values {
	my $sheet = shift;
	my $y = shift; # row
	my $col_length = shift; # 8 for alarms, 4 for status and commands
	my $return_value_col_length = shift; # 4 for alarm and status, 5 for commands

	my $noReturnValues = 0;
	my $x = $col_length; # first return value
	while(test($sheet, $y, $x)) {
		$noReturnValues++;
		$x += $return_value_col_length;
	}
	return $noReturnValues;
}
