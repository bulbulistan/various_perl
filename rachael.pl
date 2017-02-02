#!/usr/bin/perl
use warnings;
use File::Copy;
use Cwd;

my $pwd = getcwd
  ; #this assumes the script is run in the same directory as the *.wav and *.TextGrid files
my $input =
  @ARGV[0];    #supply the spreadsheet data as a CSV file (easiest solution)
my %ftb = ();

open( F, '<', $input ) || die "couldn't open the file!";

while ( $line = <F> ) {    #iterate over the CSV line by line
    @l = split( ",", $line );    #split the lines

    $filename = @l[0];           #get the filename from the first column ...
    $tobi     = @l[7];           #... and the ToBi string from the 8th

    $filename =~ s/\..*//;       #remove the file extension
    $tobi =~
      tr/\!\*\% /eaps/; #remove characters that can't be used in directory names

    $ftb{$filename} = $tobi
      ; #add filenames and ToBi strings to a hash with filenames as keys (because they're unique)

}

for my $file ( keys %ftb ) {    #now iterate over the hash
    my $newdir = $pwd . "/"
      . $ftb{$file}
      ;    #$ftb{$file} is the ToBi string to be used as the new directory name
    mkdir $newdir
      unless -d $newdir
      || $file =~ "File";    #create the new directory unless it exists

    $oldwavfile = $pwd . "/" . $file . ".wav";
    $newwavfile = $newdir . "/" . $file . ".wav";
    copy( $oldwavfile, $newwavfile );

    $oldtgfile = $pwd . "/" . $file . ".TextGrid";
    $newtgfile = $newdir . "/" . $file . ".TextGrid";

    copy( $oldtgfile, $newtgfile );

#uncomment the following lines if you want to delete the unsorted *.wav and *.TextGrid files
#unlink($oldwavfile);
#unlink($oldtgfile);

}

close(F);

