#!/usr/bin/perl

#
# A simple script to convert Excel data into
# something Quickbooks can read. 
#

use Getopt::Long;
use Text::CSV;
 
$input_file = "";
$type = "" ;
$rows = 0;
 
if ( @ARGV > 0 ) {
	    GetOptions('file=s' => \$input_file,
		       'type=s' => \$type)
	      or die ("Invalid args"); 
}

$date = time();
$output_file = "$input_file.$date.iif";

print "input file: $input_file\n";
print "type: $type\n" ;

open(INPUT, $input_file) || die("Could not open $input_file $?"); 
open(OUTPUT, ">$output_file") || die ("Could not open $output_file! $?");

sub bank () {
	# set up the document
        print OUTPUT "!TRNS\tTRNSID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
        print OUTPUT "!SPL\tSPLID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
        print OUTPUT "!ENDTRNS\n";        
        
	# get the file
	while ($line = <INPUT>) {
	     chomp ($line);
	     if ($rows == 0) { $rows++; next; }
	     @columns = split ',', $line;
	     $cramt = $columns[3] * -1;
	     
	     print OUTPUT "TRNS\t$date\tGENERAL JOURNAL\t$columns[1]\t$columns[2]\t$columns[0]\t$columns[3]\t$columns[7]\n";
	     print OUTPUT "SPL\t$date\tGENERAL JOURNAL\t$columns[1]\t$columns[5]\t$columns[0]\t$cramt\t$columns[7]\n";
	     print OUTPUT "ENDTRNS\n";
	     $rows++;
	}
        print "there are $rows rows\n";
	print OUTPUT "ENDTRNS\n";
}


sub cashxfer () {
        # set up the document
    print OUTPUT "!TRNS\tTRNSID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
    print OUTPUT "!SPL\tSPLID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
    print OUTPUT "!ENDTRNS\n";

        # get the file
    while ($line = <INPUT>) {
	chomp ($line);
	if ($rows == 0) { $rows++; next; }
	@columns = split ',', $line;
	$cramt = $columns[5] * -1;

	print OUTPUT "TRNS\t$date\tGENERAL JOURNAL\tTF Cash XFER\t$columns[1]$columns[2]\t$columns[0]\t$columns[5]\t$columns[6]\n";
	print OUTPUT "SPL\t$date\tGENERAL JOURNAL\tTF Cash XFER\t$columns[3]$columns[4]\t$columns[0]\t$cramt\t$columns[6]\n";
	print OUTPUT "ENDTRNS\n";
	$rows++;
    }
    print "there are $rows rows\n";
    print OUTPUT "ENDTRNS\n";
}

sub balance () {
	# set up the document
        print OUTPUT "!TRNS\tTRNSID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
        print OUTPUT "!SPL\tSPLID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
        print OUTPUT "!ENDTRNS\n";        
        
	# get the file
	while ($line = <INPUT>) {
	     chomp ($line);
		 @columns = split ',', $line;
		 
		 if ($rows == 0) {
			$DATE = "";
			$DATE = $columns[1];
			print OUTPUT "TRNS\t$date\tGENERAL JOURNAL\t";
			$rows++;
			next;
			}
		 if ($rows == 1) {
			$TRNSTYPE = "";
			$TRNSTYPE = $columns[1];
			print OUTPUT "$columns[1]\t"; 
			$rows++;
			next;
			}
		# blank row
		if ($rows == 2) {
			$rows++;
			next;
		}
	     
		 # header row
		 if ($rows == 3) {
			$rows++;
			next;
		}
	    if ($rows == 4 ) {
			$rows++;
			$cramt = $columns[3] * -1;
			print OUTPUT "$columns[0]\t$DATE\t$cramt\t$columns[4]\n";
			next;
		}
		
		# the rest of the data
		#$cramt = $columns[3] * -1;
		print OUTPUT "SPL\t$date\tGENERAL JOURNAL\t$TRNSTYPE\t$columns[0]$columns[1]\t$DATE\t$columns[3]\t$columns[4]\n";
	     
	}
	print OUTPUT "ENDTRNS\n";
}


sub bonus () {
	# set up the document
        print OUTPUT "!TRNS\tTRNSID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
        print OUTPUT "!SPL\tSPLID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
        print OUTPUT "!ENDTRNS\n";        
        
	# get the file
	while ($line = <INPUT>) {
	     chomp ($line);
		 @columns = split ',', $line;
		 
		 if ($rows == 0) {
			$DATE = "";
			$DATE = $columns[1];
			print OUTPUT "TRNS\t$date\tGENERAL JOURNAL\t";
			$rows++;
			next;
			}
		 if ($rows == 1) {
			$TRNSTYPE = "";
			$TRNSTYPE = $columns[1];
			print OUTPUT "$columns[1]\t"; 
			$rows++;
			next;
			}
		# blank row
		if ($rows == 2) {
			$rows++;
			next;
		}
	     
		 # header row
		 if ($rows == 3) {
			$rows++;
			next;
		}
	    if ($rows == 4 ) {
			$rows++;
			print OUTPUT "$columns[0]\t$DATE\t$columns[2]\t$columns[4]\n";
			next;
		}
		
		# the rest of the data
		$cramt = $columns[3] * -1;
		print OUTPUT "SPL\t$date\tGENERAL JOURNAL\t$TRNSTYPE\t$columns[0]$columns[1]\t$DATE\t$cramt\t$columns[4]\n";
	     
	}
	print OUTPUT "ENDTRNS\n";
}

sub custdep () {
	@data = ();
	%sums = ();
	# set up the document
    print OUTPUT "!TRNS\tTRNSID\tDOCNUM\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tMEMO\n";
    print OUTPUT "!SPL\tSPLID\tDOCNUM\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tMEMO\n";
    print OUTPUT "!ENDTRNS\n";        
        
	# get the file
	$count = 0;
	while ($line = <INPUT>) {
		chomp($line);
		# first line is the header
		if ($count == 0) { $count++; next; }
		# print OUTPUT "$line\n";
		@row = split ',', $line;
		push @data, [ @row ];	
	}

	$lines = @data;
	print "lines: $lines\n";
	
	# generate daily totals
	for (my $i = 0; $i < $lines; $i++) {
		$data[$i][3] =~ s/\s+$//;
		$sums{$data[$i][0]} += $data[$i][3]; 
		print "added $data[$i][$3] to $data[$i][0] at loc $i\n";
	}
	
	# print the iif file
	foreach my $day (sort keys %sums) {
		print "we're in the foreach loop\n";
		print "$day : $sums{$day}\n";
		
		if ($sums{$day}=="0") { print "i am a fucking fucktard\n";}
		
		print OUTPUT "TRNS\t$date\t$data[0][1]\tGENERAL JOURNAL\t$day\t$data[0][2]\t$sums{$day}\t$data[0][4]\n";
		for (my $j = 0; $j <= $lines; $j++) { 
		
			if ($data[$j][0] eq $day) {
			print "got a row to print: $day\n";
			$cramt = $data[$j][7] * -1;
			print OUTPUT "SPL\t$date\t$data[$j][1]\tGENERAL JOURNAL\t$day\t$data[$j][5]$data[$j][6]\t$cramt\t$data[$j][8]\n";
		}
		
		}
		
	print OUTPUT "ENDTRNS\n";
	}
}


sub did () {
	@data = ();
	# set up the document
    print OUTPUT "!TRNS\tTRNSID\tDOCNUM\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tMEMO\n";
    print OUTPUT "!SPL\tSPLID\tDOCNUM\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tMEMO\n";
    print OUTPUT "!ENDTRNS\n";        
        
	# get the file
	$count = 0;
	while ($line = <INPUT>) {
		# first line is the header
		if ($count == 0) { $count++; next; }
		
		chomp($line);
		# print OUTPUT "$line\n";
		@row = split ',', $line;
		push @data, [ @row ];	
		$count++;
	}

	$lines = @data;
	print "lines: $lines\n";
	
	# generate daily totals
	for (my $i = 0; $i < $lines; $i++) {
		print "working on $data[$i][0]\n";
		$sums{$data[$i][0]} += $data[$i][6]; 
	}
	
	# print the iif file
	foreach my $day (sort keys %sums) {
		#print "we're in the foreach loop\n";
		print "$day : $sums{$day}\n";
		$cramt = $sums{$day} * -1;
		print OUTPUT "TRNS\t$date\t$data[0][1]\tGENERAL JOURNAL\t$day\t$data[0][4]\t$cramt\tDID Sales\n";
		for (my $j = 0; $j <= $lines; $j++) { 
		
		if ($day >0 && $data[$j][0] eq $day) {
		#print OUTPUT "got a row to print: $day\n";
		$data[$j][6] =~ s/\s+$//;
		print OUTPUT "SPL\t$date\t$data[$j][1]\tGENERAL JOURNAL\t$day\t$data[$j][2]$data[$j][3]\t$data[$j][6]\t$data[$j][5]\n";
		}
		
		}
	print OUTPUT "ENDTRNS\n";
	}
}   

sub fee () {
    # set up the document
    print OUTPUT "!TRNS\tTRNSID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
    print OUTPUT "!SPL\tSPLID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
    print OUTPUT "!ENDTRNS\n";

    # get the file
    while ($line = <INPUT>) {
	chomp ($line);
	 # first row is header data
	if ($rows == 0) { $rows++; next; }
	@columns = split ',', $line;
	$cramt = $columns[4] * -1;

	print OUTPUT "TRNS\t$date\tGENERAL JOURNAL\t$columns[1]\t$columns[2]$columns[3]\t$columns[0]\t$columns[4]\t$columns[8]\n";
	print OUTPUT "SPL\t$date\tGENERAL JOURNAL\t$columns[1]\t$columns[6]\t$columns[0]\t$cramt\t$columns[8]\n";
	print OUTPUT "ENDTRNS\n";
	$rows++;
    }
    print "there are $rows rows\n";
    print OUTPUT "ENDTRNS\n";
}


sub refund () {
	@data = ();
	# set up the document
    print OUTPUT "!TRNS\tTRNSID\tDOCNUM\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tMEMO\n";
    print OUTPUT "!SPL\tSPLID\tDOCNUM\tTRNSTYPE\tDATE\tACCNT\tAMOUNT\tMEMO\n";
    print OUTPUT "!ENDTRNS\n";        
        
	# get the file
	
	while ($line = <INPUT>) {
		 # first row is header data
		if ($rows == 0) { $rows++; next; }
		chomp($line);
		# print OUTPUT "$line\n";
		@row = split ',', $line;
		push @data, [ @row ];	
	}

	$lines = @data;
	print "lines: $lines\n";
	
	# generate daily totals
	for (my $i = 0; $i < $lines; ++$i) {
		$sums{$data[$i][0]} += $data[$i][6]; 
	}
	
	# print the iif file
	foreach my $day (sort keys %sums) {
		#print "we're in the foreach loop\n";
		#print "$day : $sums{$day}\n";
		
		print OUTPUT "TRNS\t$date\t$data[0][1]\tGENERAL JOURNAL\t$day\t$data[0][2]\t$sums{$day}\t$data[0][5]\n";
		for (my $j = 0; $j <= $lines; $j++) { 
		
		if ($data[$j][0] eq $day) {
		#print OUTPUT "got a row to print: $day\n";
		$cramt = $data[$j][6] * -1;
		print OUTPUT "SPL\t$date\t$data[$j][1]\tGENERAL JOURNAL\t$day\t$data[$j][3]$data[$j][4]\t$cramt\t$data[$j][5]\n";
		}
		
		}
	print OUTPUT "ENDTRNS\n";
	}
}



sub revenue () {
	# set up the document
        print OUTPUT "!TRNS\tTRNSID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
        print OUTPUT "!SPL\tSPLID\tTRNSTYPE\tDOCNUM\tACCNT\tDATE\tAMOUNT\tMEMO\n";
        print OUTPUT "!ENDTRNS\n";        
        
	# get the file
	while ($line = <INPUT>) {
	     chomp ($line);
		 @columns = split ',', $line;
		 
		 if ($rows == 0) {
			$DATE = "";
			$DATE = $columns[1];
			print OUTPUT "TRNS\t$date\tGENERAL JOURNAL\t";
			$rows++;
			next;
			}
		 if ($rows == 1) {
			$TRNSTYPE = "";
			$TRNSTYPE = $columns[1];
			print OUTPUT "$columns[1]\t"; 
			$rows++;
			next;
			}
		# blank row
		if ($rows == 2) {
			$rows++;
			next;
		}
	     
		 # header row
		 if ($rows == 3) {
			$rows++;
			next;
		}
	     if ($rows == 4){
			$rows++;
			$cramt = $columns[3] * -1;
			print OUTPUT "$columns[0]\t$DATE\t$cramt\t$columns[4]\n";
			next;
		}
		
	     if ($rows == 5 || $rows == 6 ){
		 $rows++; 
		 $cramt = $columns[3] * -1;
		 print OUTPUT "SPL\t$date\tGENERAL JOURNAL\t$TRNSTYPE\t$columns[0]$columns[1]\t$DATE\t$cramt\t$columns[4]\n";
		 next;
	     }

		# the rest of the data
		print OUTPUT "SPL\t$date\tGENERAL JOURNAL\t$TRNSTYPE\t$columns[0]$columns[1]\t$DATE\t$columns[2]\t$columns[4]\n";
		$rows++;
	     
	}
	print OUTPUT "ENDTRNS\n";
}

if ($type eq "bank") {&bank()};
if ($type eq "bonus") {&bonus()};
if ($type eq "custdep") {&custdep()};
if ($type eq "did") {&did()};
if ($type eq "fee") {&fee()};
if ($type eq "refund") {&refund()};
if ($type eq "revenue") {&revenue()};
if ($type eq "cashxfer") {&cashxfer()};
if ($type eq "balance") {&balance()};

close (OUTPUT);
close (INPUT);
system ("/usr/bin/unix2dos", $output_file);

		
		
		

	
