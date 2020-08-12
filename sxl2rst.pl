#!/usr/bin/perl -w
# sxl2rst.pl [SXL in xlsx format]
# Convert signal exchange list (SXL) for RSMP in xlsx or csv format to ReStructuredText
# Requires Spreadsheet::XLSX (sudo apt-get install libspreadsheet-xlsx-perl)
# and python3-tabulate for generating tables

use strict;
use Spreadsheet::XLSX;
use Getopt::Long;

my $omit_objects;
my $omit_object_col;
my $omit_xnid_col;
my $omit_xnacid_col;
my $csv;
my $help;
GetOptions(
	# Omits object sheet
	"omit-objects" => \$omit_objects,

	# Omits object column in alarms, status and commands
	"omit-object-col" => \$omit_object_col,

	# Omits externalAlarmCodeId (xNid) column
	"omit-xnid-col" => \$omit_xnid_col,

	# Omits externalNtsAlarmCodeId (xNACId) column
	"omit-xnacid-col" => \$omit_xnacid_col,

	# Read SXL as CSV-files (zipped) instead of Excel-file
	"csv" => \$csv,

	"h|help" => \$help,
);

if(defined($help)) {
	die("usage sxl2rst.pl [--omit-objects] [--omit-object-col] [--omit-xnid-col] [--omit-xnacid-col] [--csv] [FILE]");
}

my @files = @ARGV;
my $fname;
my $sheet;

foreach $fname (@files) {
	read_sxl($fname) unless(defined($csv));
	read_csv($fname) if(defined($csv));
}
rst_line_break_substitution();


# Read SXL in Excel format
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
			print_objects($sheet) unless(defined($omit_objects));
		}
	}
}

# Read SXL in CSV format
sub read_csv {
	my $fname = shift;

	my $sheet_v;
	my $sheet_ot;
	my $sheet_agg;
	my $sheet_a;
	my $sheet_s;
	my $sheet_c;
	my $sheet_o;

	system("unzip -q $fname");
	my @cfile = `ls *.csv`;
	foreach my $file (@cfile) {
		chomp $file;
		open(my $data, '<', $file) or die;

		# Read csv file
		my $line_break = 0; # line breaks has to be handled
		my $y = 0;
		my $x = 0; 
		while(my $line = <$data>) {
			chomp $line;

			my @fields = split ";", $line;

			# This implementations has some problems dealing with
			# quotations within CSV files. Excel quotes any cells
			# containing illegal characters, e.g. linesbreaks and
			# semicolon. Currently we treat everything as linebreak

			$x = 0 if($line_break == 0);
			foreach my $field (@fields) {
				if($line_break == 1) {
					$sheet->{Cells}[$y][$x]->{Val} = "$sheet->{Cells}[$y][$x]->{Val}\r\n$field";
				} else {
					$sheet->{Cells}[$y][$x]->{Val} = $field;
				}

				if($field =~ /"/) {
					$sheet->{Cells}[$y][$x]->{Val} =~ s/"//;
					if($line_break == 1) {
						$line_break = 0;
						$x++;
					} else {
						$line_break = 1;
					}
				} else {
					if($line_break == 0) {
						$x++;
					}
				}

			}
			$y++ if($line_break == 0);
		}

		# Check which sheet by checking values
		if($sheet->{Cells}[1][1]->{Val} eq "Signal Exchange List") {
			$sheet_v = $sheet;
		} elsif($sheet->{Cells}[0][0]->{Val} eq "Object types") {
			$sheet_ot = $sheet;
		} elsif($sheet->{Cells}[0][0]->{Val} eq "Aggregated status per grouped object") {
			$sheet_agg = $sheet;
		} elsif($sheet->{Cells}[0][0]->{Val} eq "Alarms per object type") {
			$sheet_a = $sheet;
		} elsif($sheet->{Cells}[0][0]->{Val} eq "Status per object type") {
			$sheet_s = $sheet;
		} elsif($sheet->{Cells}[0][0]->{Val} eq "Commands per object type") {
			$sheet_c = $sheet;
		} else {
			$sheet_o = $sheet;
		}
		$sheet = undef;
	}
	print_version($sheet_v);		# Version
	print_object_types($sheet_ot);		# Object types
	print_objects($sheet_o) unless(defined($omit_objects)); # Objects
	print_aggregated_status($sheet_agg);	# Aggregated status
	print_alarms($sheet_a);
	print_status($sheet_s);
	print_commands($sheet_c);

	system("rm *.csv");
}

sub print_version {
	my $sheet = shift;
	printf("Signal Exchange List\n");
	printf("====================\n");

	cprint($sheet, 3, 1, "Plant Id");
	cprint($sheet, 5, 1, "Plant Name");
	cprint($sheet, 9, 1, "Constructor");
	cprint($sheet, 11,1, "Reviewed");
	cprint($sheet, 14,1, "Approved");
	cprint($sheet, 17,1, "Created date");
	cprint($sheet, 20,1, "SXL revision");
	cprint($sheet, 20,2, "Revision date");
	cprint($sheet, 25,1, "RSMP version");
}

sub print_object_types {
	my $sheet = shift;
	my $y;
	my $yf;
	# Object types sheet
	printf "\n";
	printf("Object Types\n");
	printf("------------\n");

	# Print all grouped objects
	printf "\n";
	printf "Grouped objects\n";
	printf "^^^^^^^^^^^^^^^\n";
	printf "\n";

	printf ".. figtable::\n";
   	printf "   :nofig:\n";
        printf "   :label: Grouped objects\n";
        printf "   :caption: Grouped objects\n";
        printf "   :loc: H\n";
        printf "   :spec: >{\\raggedright\\arraybackslash}";
	printf "p{0.30\\linewidth} ";
        printf "p{0.50\\linewidth}\n";
	printf "\n";
	printf "   ==========     ===========\n";
	printf "   ObjectType     Description\n";
	printf "   ==========     ===========\n";
	$y = 6;
	while (test($sheet, $y, 0)) {
		$yf = sprintf("%03d", $y);
		printf "   |go-o$yf|      |go-d$yf|\n";
		$y++;
	}
	printf "   ==========     ===========\n\n";
	printf "..\n\n";
	$y = 6;
	while (test($sheet, $y, 0)) {
		otprint($sheet, "go", $y);
		$y++;
	}

	# Print all single objects
	printf "\n";
	printf "Single objects\n";
	printf "^^^^^^^^^^^^^^\n";
	printf "\n";

	printf ".. figtable::\n";
   	printf "   :nofig:\n";
        printf "   :label: Single objects\n";
        printf "   :caption: Single objects\n";
        printf "   :loc: H\n";
        printf "   :spec: >{\\raggedright\\arraybackslash}";
	printf "p{0.30\\linewidth} ";
        printf "p{0.50\\linewidth}\n";
	printf "\n";
	printf "   ==========     ===========\n";
	printf "   ObjectType     Description\n";
	printf "   ==========     ===========\n";
	$y = 18;
	while (test($sheet, $y, 0)) {
		$yf = sprintf("%03d", $y);
		printf "   |so-o$yf|      |so-d$yf|\n";
		$y++;
	}
	printf "   ==========     ===========\n\n";
	printf "..\n\n";
	$y = 18;
	while (test($sheet, $y, 0)) {
		otprint($sheet, "so", $y);
		$y++;
	}
}

sub print_aggregated_status {
	my $sheet = shift;
	my $y;
	my $yf;
	# Aggregated status sheet
	printf "\n";
	printf("Aggregated status\n");
	printf("-----------------\n");
	printf "\n";

	# Print aggregated status
	printf ".. figtable::\n";
   	printf "   :nofig:\n";
        printf "   :label: Aggregated status\n";
        printf "   :caption: Aggregated status\n";
        printf "   :loc: H\n";
        printf "   :spec: >{\\raggedright\\arraybackslash}";
	printf "p{0.15\\linewidth} ";
        printf "p{0.20\\linewidth}  ";
        printf "p{0.18\\linewidth}  ";
        printf "p{0.18\\linewidth}  ";
        printf "p{0.15\\linewidth}\n";
	printf "\n";
	printf "   ==========     ===========  ==================  =============== ===========\n";
	printf "   ObjectType     Status       functionalPosition  functionalState Description\n";
	printf "   ==========     ===========  ==================  =============== ===========\n";
	$y = 6;
	while (test($sheet, $y, 0)) {
		$yf = sprintf("%03d", $y);
		printf "   |ag-1$yf|      |ag-2$yf|    |ag-3$yf|           |ag-4$yf|       |ag-5$yf|\n";
		$y++;
	}
	printf "   ==========     ===========  ==================  =============== ===========\n\n";
	printf "..\n\n";
	$y = 6;
	while (test($sheet, $y, 0)) {
		aggprint($sheet, $y);
		$y++;
	}

	# Print all state bits
	printf "\n";
	printf ".. figtable::\n";
   	printf "   :nofig:\n";
        printf "   :label: State bits\n";
        printf "   :caption: State bits\n";
        printf "   :loc: H\n";
        printf "   :spec: >{\\raggedright\\arraybackslash}";
	printf "p{0.15\\linewidth} ";
        printf "p{0.30\\linewidth}  ";
        printf "p{0.45\\linewidth}\n";
	printf "\n";
	printf "   =============  ===========================  ========\n";
	printf "   |statebit|     Description                  Comment\n";
	printf "   =============  ===========================  ========\n";
	stateprint($sheet);
	printf "   =============  ===========================  ========\n";
	printf "..\n\n";
	print ".. |statebit| replace:: State- Bit nr (1234567)\n\n";

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
	printf "\n";
	printf "Grouped objects\n";
	printf "---------------\n";
	printf "\n";
	printf ".. list-table:: Grouped objects\n";
	printf "   :widths: 10 10 10 10 10 20\n";
	printf "   :header-rows: 1\n";
	printf "\n";
	printf "   * - ObjectType\n";
	printf "     - Object\n";
	printf "     - componentId\n";
	printf "     - NTSObjectId\n";
	printf "     - externalNtsId\n";
	printf "     - Description\n";
	my $y = 6;
	while (test($sheet, $y, 0)) {
		oprint($sheet, $y);
		$y++;
	}

	# Print all single objects
	printf "\n";
	printf "Single objects\n";
	printf "--------------\n";
	printf "\n";
	printf ".. list-table:: Single objects\n";
	printf "   :widths: 10 10 10 10 10 20\n";
	printf "   :header-rows: 1\n";
	printf "\n";
	printf "   * - ObjectType\n";
	printf "     - Object\n";
	printf "     - componentId\n";
	printf "     - NTSObjectId\n";
	printf "     - externalNtsId\n";
	printf "     - Description\n";
	$y = 24;
	while (test($sheet, $y, 0)) {
		oprint($sheet, $y);
		$y++;
	}
}

sub print_alarms {
	my $sheet = shift;
	printf "\n";
	printf("Alarms\n");
	#printf("======\n");
	printf("------\n");

	# Print header
	my @widths = ();
	my @table_headers = ();
	push(@table_headers, "ObjectType"); push @widths, "0.15";
	unless(defined($omit_object_col)) {
		push(@table_headers, "Object (optional)");
		push(@widths, "0.10");
	}
	push(@table_headers, "alarmCodeId"); push @widths, "0.10";
	unless(defined($omit_xnid_col)) {
		push(@table_headers, "externalAlarmCodeId");
		push(@widths, "0.10");
	}
	unless(defined($omit_xnacid_col)) {
	        push(@table_headers, "externalNtsAlarmCodeId");
		push(@widths, "0.10");
	}
	push(@table_headers, ("Description", "Priority", "Category"));
	push(@widths, ("0.45", "0.07", "0.07"));

	my ($fh, $file) = start_figtable(\@widths, \@table_headers, "Alarms");

	# Print alarms
	my $y = 6;
	while (test($sheet, $y, 7)) {
		my $has_return_values = get_no_return_values($sheet, $y, 8, 4);
		my @table_data = aprint($sheet, $y, 8, $has_return_values);
		add_figtable($fh, \@table_data);
		$y++;
	}
	end_figtable($fh, $file);

	$y = 6;
	while (test($sheet, $y, 7)) {
		if(get_no_return_values($sheet, $y, 8, 4) > 0) {
			my $xCodeId = $sheet->{Cells}[$y][2]->{Val};
			
			# Print header
			print "\n";
			print "$xCodeId\n";
			print "^^^^^\n\n";

			print $sheet->{Cells}[$y][3]->{Val}."\n\n"; # Description

			@widths = ("0.15", "0.08", "0.13", "0.35");
			@table_headers = ("Name", "Type", "Value", "Comment");
			($fh, $file) = start_figtable(\@widths, \@table_headers, $xCodeId);
			rprint($sheet, $fh, $y, 8, 4, 2, 3);
			end_figtable($fh, $file);
		}
		$y++
	}
}

sub print_status {
	my $sheet = shift;
	printf("\n");
	printf("Status\n");
	#printf("======\n");
	printf("------\n");
	printf("\n");

	# The status table can be so long it won't fit on a single page
	print ".. raw:: latex\n\n";
	print "    \\newpage\n\n";

	# Print header
	my @widths = ();
	my @table_headers = ();

	push(@table_headers, "ObjectType"); push @widths, "0.24";
	unless(defined($omit_object_col)) {
		push(@table_headers, "Object (optional)");
		push(@widths, "0.10");
	}
	push(@table_headers, "statusCodeId"); push @widths, "0.10";
	push(@table_headers, "Description");
	push(@widths, "0.55");


	my ($fh, $file) = start_figtable(\@widths, \@table_headers, "Status");
	
	# Print status
	my $y = 6;
	while (test($sheet, $y, 7)) {
		my $has_return_values = get_no_return_values($sheet, $y, 4, 4);
		my @table_data = aprint($sheet, $y, 4, $has_return_values);
		add_figtable($fh, \@table_data);
		$y++;
	}
	end_figtable($fh, $file);

	# Print return values

	$y = 6;
	while (test($sheet, $y, 7)) {
		if(get_no_return_values($sheet, $y, 4, 4) > 0 ) {
			my $xCodeId = $sheet->{Cells}[$y][2]->{Val};

			# Print header
			print "\n";
			print "$xCodeId\n";
			print "^^^^^^^^\n";

			print "\n";
			print $sheet->{Cells}[$y][3]->{Val}."\n\n"; # Description

			@widths = ("0.15", "0.08", "0.13", "0.50");
			@table_headers = ("Name", "Type", "Value", "Comment");
			($fh, $file) = start_figtable(\@widths, \@table_headers, $xCodeId);
			rprint($sheet, $fh, $y, 4, 4, 2, 3);
			end_figtable($fh, $file);
		}
		$y++;
	}
}

sub print_commands {
	my $sheet = shift;
	printf "\n";
	printf "Commands\n";
	#printf "========\n";
	printf "--------\n";

	my $sec;
	my $y;
	my @sections = command_section($sheet);

	# Print header
	my @widths = ();
	my @table_headers = ();
	push(@table_headers, "ObjectType"); push @widths, "0.24";
	unless(defined($omit_object_col)) {
		push(@table_headers, "Object (optional)");
		push(@widths, "0.10");
	}
	push(@table_headers, "commandCodeId"); push @widths, "0.15";
	push(@table_headers, "Description"); push(@widths, "0.40");

	my ($fh, $file) = start_figtable(\@widths, \@table_headers, "Commands");

	foreach $sec (@sections) {
		# Need to check each command section
		$y = $sec;

		# Print command
		while (test($sheet, $y, 0)) {
			my $has_return_values = get_no_return_values($sheet, $y, 4, 5);
			my @table_data = aprint($sheet, $y, 4, $has_return_values);
			add_figtable($fh, \@table_data);
			$y++;
		}
	}
	end_figtable($fh, $file);

	# Print arguments
	foreach $sec (@sections) {
		# Need to check each command section
		$y = $sec;
		# Print command
		while (test($sheet, $y, 0)) {
			if(get_no_return_values($sheet, $y, 4, 5) > 0) {
				my $xCodeId = $sheet->{Cells}[$y][2]->{Val};
				my $description = $sheet->{Cells}[$y][3]->{Val};

				# Fix line breaks
				$description = rst_line_breaks($description, "\n");

				# Print header
				print "\n";
				print "$xCodeId\n";
				print "^^^^^\n";
				print "\n";

				print "$description\n\n";

				@widths = ("0.14", "0.20", "0.07", "0.15", "0.30");
				@table_headers = ("Name", "Command", "Type", "Value", "Comment");
				($fh, $file) = start_figtable(\@widths, \@table_headers, $xCodeId);
				rprint($sheet, $fh, $y, 4, 5, 3, 4);
				end_figtable($fh, $file);

				#$return_text .= rprint($sheet, $y, 4, 5, 3, 4, $txt);

			}
			$y++;
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

	printf "+ **$text**:";
        $val =~ s/%/%%/g if(defined($val)); # Needed for printf()
	printf " $val" if(defined($val) && ($val ne ""));
	printf "\n";
}

# Print object type
sub otprint {
	my $sheet = shift;
	my $object_type = shift; #go = grouped objects, so = single objects
	my $y = shift;
	my $yf = sprintf("%03d", $y);

	my $objecttype = $sheet->{Cells}[$y][0]->{Val};
	my $description = $sheet->{Cells}[$y][5]->{Val};
	$description = "--" unless(defined($description));

	printf ".. |$object_type-o$yf| replace:: $objecttype\n\n";
	printf ".. |$object_type-d$yf| replace:: $description\n\n";
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
		printf "   * - $objecttype\n";
		printf "     - $object\n";
		printf "     - $cId\n";
		printf "     - $ntsoId\n";

		# ExternalNtsId
		printf "     - $externalNtsId\n" if(defined($externalNtsId));

		printf "     - $description\n" if(defined($description));
	}
}

# Print aggregated status
sub aggprint {
	my $sheet = shift;
	my $y = shift;
	my $yf = sprintf("%03d", $y);

	my $objecttype = $sheet->{Cells}[$y][0]->{Val};
	my $state = $sheet->{Cells}[$y][1]->{Val};
	my $functionalPosition = $sheet->{Cells}[$y][2]->{Val};
	my $functionalState = $sheet->{Cells}[$y][3]->{Val};
	my $description = $sheet->{Cells}[$y][4]->{Val};
	$state = "--" unless(defined($state));
	$functionalPosition = "--" unless(defined($functionalPosition));
	$functionalState = "--" unless(defined($functionalState));
	$description = "--" unless(defined($description));

	print ".. |ag-1$yf| replace:: $objecttype\n\n";
	print ".. |ag-2$yf| replace:: $state\n\n";
	print ".. |ag-3$yf| replace:: $functionalPosition\n\n";
	print ".. |ag-4$yf| replace:: $functionalState\n\n";
	print ".. |ag-5$yf| replace:: $description\n\n";
}

# Print state bits (aggregated status)
sub stateprint {
	my $sheet = shift;
	my $y;
	my @bit;
	for($y = 0; $y < 8; $y++) {
		$bit[$y] = $sheet->{Cells}[$y+16][2]->{Val};
		$bit[$y] = "--" unless(defined($bit[$y]));

		# Remove line breaks
		$bit[$y] =~ s/\r\n/ /g;
	}
	printf "   1              Local mode                   $bit[0]\n";
	printf "   2              No communications            $bit[1]\n";
	printf "   3              High priority fault          $bit[2]\n";
	printf "   4              Medium priority fault        $bit[3]\n";
	printf "   5              Low priority fault           $bit[4]\n";
	printf "   6              Connected / Normal - In Use  $bit[5]\n";
	printf "   7              Connected / Normal - Idle    $bit[6]\n";
	printf "   8              Not Connected                $bit[7]\n";
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
		$val[$i] = "" unless(defined($val[$i]));
	}
	
	# Make column into a link if return values exist
	if($has_return_values > 0) {
		$val[$col_link] = "`".$val[$col_link]."`_"; 
	}

	# Print row
	my @row;
	for($i = 0; $i < $col_length; $i++) {

		# Skip object column, if set
		if(($i == 1) && defined($omit_object_col)) {
			$i++;
		}

		# Skip externalAlarmCodeId column, if set
		if(($i == 4) && defined($omit_xnid_col)) {
			$i++
		}

		# Skip externalNtsAlarmCodeId column, if set
		if(($i == 5) && defined($omit_xnacid_col)) {
			$i++
		}

		$val[$i] = rst_line_breaks($val[$i], "\n");
		push @row, $val[$i];
	}
	return @row;
}

# Print arguments/return values
sub rprint {
	my $sheet = shift;
	my $fh = shift;
	my $y = shift;  # Start row
	my $start_x = shift; # 8 for alarms, 4 for status and commands
	my $return_value_col_length = shift; # 4 for alarm and status, 5 for commands
	my $value_list_col = shift; # this column of return values/arguments should be split into bullet list, 2 for alarm and status, 3 for commands
	my $comment_list_col = shift; # this column of return values/arguments should be split into bullet list

	my $i;
	my @val;
	my $x = $start_x;
	my $col_length;
	my @data;

	# return values
	while(test($sheet, $y, $x)) {
		# Get values for a row
		$col_length = $return_value_col_length;
		for($i = 0; $i < $col_length; $i++) {
			$val[$i] = $sheet->{Cells}[$y][$x++]->{Val};
			$val[$i] = "" unless(defined($val[$i]));
		}

		# Check for semicolon in the "Comment" field, which is the last one
		semi_check($sheet, $x, $y);

		# Print row
		@data = ();
		for($i = 0; $i < $col_length; $i++) {
			$val[$i] = rst_line_breaks($val[$i], "\n");
			push(@data, $val[$i]);
		}
		add_figtable($fh, \@data);
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
	if(defined($comment)) {
		printf STDERR "WARNING: Found semicolon in comment field: $comment\n" if ($comment =~ /;/);
	}
}

# Test for contens if the first column
sub test {
	my $sheet = shift;
	my $y = shift;
	my $x = shift;

	# If using CSV-files as source, treat empty values as undef
	if(defined($sheet->{Cells}[$y][$x]->{Val})) {
		if(($sheet->{Cells}[$y][$x]->{Val} eq "") || ($sheet->{Cells}[$y][$x]->{Val} eq "\r")) {
			$sheet->{Cells}[$y][$x]->{Val} = undef;
		}
	}
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

# Insert ReStructuredText line breaks
sub rst_line_breaks {
	my $txt = shift;
	my $linebreak = shift; # Usually \n; 

	$txt =~ s/\r//g;
	$txt =~ s/$linebreak/ |br| /g;
	$txt =~ s/\|br\| $//g; # Trailing line break

	return $txt;
}

# Normally it is not possible to insert line breaks in Sphinx
# It is possible to use "line bock" (|), but it also adds another
# line break in the beginning and ending which is not pretty
# inside a table
sub rst_line_break_substitution {
	print "\n";
	print ".. |br| replace:: |br_html| |br_latex|\n\n";
	print ".. |br_html| raw:: html\n\n";
	print "   <br>\n\n";
	print ".. |br_latex| raw:: latex\n\n";
	print "   \\newline\n\n";
}


sub start_figtable {
	my $widths = shift;
	my $headers = shift;
	my $label = shift;

	printf "\n";
	printf ".. figtable::\n";
   	printf "   :nofig:\n";
        printf "   :label: $label\n";
        printf "   :caption: $label\n";
        printf "   :loc: H\n";
        printf "   :spec: >{\\raggedright\\arraybackslash}";
	foreach (@$widths) {
		printf "p{$_\\linewidth} ";
	}
	printf "\n\n";

	use File::Temp qw(tempdir);
	my $dir = tempdir( CLEANUP => 1 );
	my $table_file = "$dir/table.txt";
	open my $fh, '>', $table_file or die;
	print $fh join(';', @$headers);
	print $fh "\n";

	return $fh, $table_file;
}

sub add_figtable {
	my $fh = shift;
	my $data = shift;

	print $fh join(';', @$data);
	print $fh "\n";
}

sub end_figtable {
	my $fh = shift;
	my $file = shift;
	
	# Use tabulate to generate table
	close $fh;
	my @output = qx{tabulate -1 -s ';' -f rst < $file};
	foreach (@output) {
		$_ =~ s/^/   /g;  # Adjust border
		print $_;
	}

	printf "..\n";
}

