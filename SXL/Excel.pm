# A perl module to read RSMP SXL in Excel format

package SXL::Excel;
use Spreadsheet::XLSX;
use strict;

sub new {
	my $class = shift;
	my $self = {
		fname => shift
	};
	bless $self, $class;
	return $self;
}

sub load {
	my $self = shift;
	return Spreadsheet::XLSX->new($self->{fname});
}
 
# Get number of arguments/return values for given row
sub get_no_return_values {
	my $self = shift;
	my $sheet = shift;
	my $y = shift; # row
	my $col_length = shift; # 8 for alarms, 4 for status and commands
	my $return_value_col_length = shift; # 4 for alarm and status, 5 for commands

	my $noReturnValues = 0;
	my $x = $col_length; # first return value
	while(test($self, $sheet, $y, $x)) {
		$noReturnValues++;
		$x += $return_value_col_length;
	}
	return $noReturnValues;
}

# Test for contens if the first column
sub test {
	my $self = shift;
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

# Find command section
# Return a list of y-positions for the start of each section
sub command_section {
	my $self = shift;
	my $sheet = shift;
	my $y = 4; # Section won't start before row 4
	my @list;
	my $text;
	while ($y<100) {
		$text = "";
		$text = $sheet->{Cells}[$y][0]->{Val} if(test($self, $sheet, $y, 0));

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

1;
