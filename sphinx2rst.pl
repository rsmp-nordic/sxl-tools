#!/usr/bin/perl -w
# sphinx2rst.pl
# Convert sphinx sources to rst format suitable for github
#
# TODO: :num:_table links does not work

# Remember number of whitespaces used in tables
my $num_ws=0;

# Source
my $dir = "/home/i0davla/rsmp/rsmp_core/sphinx/source/";
my $file = "index.rst";

# Output
open(my $out, '>', "/home/i0davla/rsmp/rsmp_core/rst/rsmp.rst") or die;

print $out "Contens\n";
print $out "=======\n\n";
read_index($file);
print $out "\n";
read_rst($file);
generate_images();

sub read_index {
	my $file = shift;
	open(my $data, '<', "$dir/$file") or die;
	my $dir=0;

	while(my $line = <$data>) {
		chomp $line;

		# Begin directives 
		if($line =~ /^\.\./) {
			if($line =~ /.. raw:: latex/) {
				# raw
				$dir=0;
			}
			if($line eq ".. toctree::") {
				$dir=2;
			}
		}
		else {
			# end directives
			if(($dir >= 1) && ($dir <= 4)) {
				unless(($line eq "") || ($line =~ /^ /)) {
					$dir=0;
				}
			}
		}

		# Print contens
		if($dir == 1) {
			# Don't print
		}
		elsif($dir == 2) {
			print_contens($line);
		}
		else {
			#print
		}
	}
	close($data);
}

sub print_contens {
	my $line = shift;

	# Remove leading whitespace
	$line =~ s/^ *//;

	# Delete toctree directives
	my $d=0;
	$d=1 if($line =~ s/..toctree:://);
	$d=1 if($line =~ s/:maxdepth://);
	$d=1 if($line =~ s/:caption://);
	$d=1 if($line =~ s/:numbered://);
	$d=1 if($line eq "");

	# Print contens
	if($d==0) {
		printf($out "* ");
		if($line =~ /\//) {
			$line =~ s/\//\/`/;
		}
		else {
			print $out "`";
		}
		print $out $line;
		print $out "`_\n";
	}
	$line =~ s/`//g;

	# Recursive search
	unless($line =~ /.rst$/) {
		$line = "$line.rst";
	}
	read_index($line) if($d==0);
}

sub read_rst {
	my $file = shift;
	open(my $data, '<', "$dir/$file") or die;
	my $dir=0;

	while(my $line = <$data>) {
		chomp $line;

		# Begin directives 
		if($line =~ /^\.\./) {
			$dir=1;
			if($line =~ /.. code/) {
				# Leave code untouched
				$dir=0;
			}
			if($line =~ /.. _/) {
				# link
				$dir=0;
			}
			if($line =~ /.. image/) {
				# fix link
				$line =~ s/msc\///;
				$line =~ s/\///;
				$dir=0;
			}
			if($line =~ /.. figure/) {
				# fix link
				$line =~ s/msc//;
				$line =~ s/\///;
				$dir=0;
			}
			if($line eq ".. toctree::") {
				$dir=2;
			}
			if($line eq ".. figtable::") {
				$dir=3;
			}
			if($line eq ".. glossary::") {
				$dir=4;
			}
		}
		else {
			# end directives
			if(($dir >= 1) && ($dir <= 4)) {
				unless(($line eq "") || ($line =~ /^ /)) {
					$dir=0;
				}
			}
		}
		
		# Convert links
		$line =~ s/:ref:\`//g;
		$line =~ s/\`/_/g;

		# Print contens
		if($dir == 1) {
			# Don't print
		}
		elsif($dir == 2) {
			print_toctree($line);
		}
		elsif($dir == 3) {
			print_figtable($line);
		}
		elsif($dir == 4) {
			print_glossary($line);
		}
		else {
			print $out "$line\n";
		}
	}
}

sub print_glossary {
	my $line = shift;

	# Delete glossary directives
	my $d=0;
	$d=1 if($line =~ s/.. glossary:://);

	# Remove leading whitespace
	$line =~ s/^   //;

	if($line eq "") {

	}
	else {
		unless($line =~ s/^   //) {
			$line = "**$line**";
		}
	}

	printf($out "$line\n") if($d==0);
}

sub print_toctree {
	my $line = shift;

	# Remove leading whitespace
	$line =~ s/^ *//;

	# Delete toctree directives
	my $d=0;
	$d=1 if($line =~ s/..toctree:://);
	$d=1 if($line =~ s/:maxdepth://);
	$d=1 if($line =~ s/:caption://);
	$d=1 if($line =~ s/:numbered://);
	$d=1 if($line eq "");

	unless($line =~ /.rst$/) {
		$line = "$line.rst";
	}
	printf($out "\n") if($d==0);
	read_rst($line) if($d==0);
}

sub print_figtable {
	my $line = shift;	

	# Remove leading whitespace
	if($line =~ /(.*):nofig:/) {
		$num_ws = $1;
	}
	$line =~ s/$num_ws//;
	#$line =~ s/^ *//;

	# Delete figtable directives
	my $d=0;
	$d=1 if($line =~ s/.. figtable:://);
	$d=1 if($line =~ s/:nofig://);
	$d=1 if($line =~ s/:label:.*//);
	$d=1 if($line =~ s/:caption:.*//);
	$d=1 if($line =~ s/:loc:.*//);
	$d=1 if($line =~ s/:spec:.*//);

	# Don't print directive (d=1)
	printf($out "$line\n") if($d==0);
}

sub generate_images {
	system("rm /home/i0davla/rsmp/rsmp_core/rst/img/*");
	system("make -C /home/i0davla/rsmp/rsmp_core/sphinx clean");
	system("make -C /home/i0davla/rsmp/rsmp_core/sphinx generated-images");
	system("cp /home/i0davla/rsmp/rsmp_core/sphinx/source/img/msc/*.png /home/i0davla/rsmp/rsmp_core/rst/img");
}
