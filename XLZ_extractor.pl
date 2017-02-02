# Create TM from Idiom v0.9
# bulbul
#
# PURPOSE:
# This is a script to create a TM from a translated WorldServer Workbench file (.wlz).
#
# OUTPUT:
# A bilingual file named "bi.txt" with source and target strings separated by tabs. 
# It can be processed in Oliphant (http://okapi.sourceforge.net/applications.html) to create a TMX.
#
# CHANGELOG: 
# 22.11.2005
# first attempts
# v0.5 - first working version
#
# 22.11.2005
# v0.6 - added segment number check and message for user, the output is now in UTF8
# 
# 25.11.2010
# v0.9 - added functionality to unpack XLZ files, rewrote the module extracting segments using LibXML, combined three scripts into one
#
# TO DO:
# - rewrote using functions
# - add output to TMX
# - add error handling and user interaction
#


use XML::LibXML;
use Archive::Zip qw( :ERROR_CODES );
use File::Glob ':glob';
use Cwd;

# 0. Prepare all the stuff

@zipy =  glob("*.xlz"); # all XLZ files in the current directory
$pwd = getcwd; # get current working directory

# 1. Extract the xml/xlf files from xlz files into the current directory

for ($i = 0; $i < scalar(@zipy); $i++) {
	my $zip = Archive::Zip->new();
	$zip->read(@zipy[$i]);
	
	@inside = $zip->memberNames();
	$membername = @inside[0];
	$newname = @zipy[$i] . "_" . $membername . "\n"; 

	$zip->extractMember($membername, $i . ".xlf");
}

# 2. Process the xml/xlf files in the current directory

my $parser = XML::LibXML->new(); # Creates a new instance of the LibXML parser

@fajly =  glob("*.x[ml][lf]"); # all XLF files in the current directory

@source = ();
@target = ();

for ($i = 0; $i < scalar(@fajly); $i++){

$doc = $parser->parse_file(@fajly[$i]); # The parser parses the XLF file

@source_node_list = $doc->getElementsByTagName('source');
@target_node_list = $doc->getElementsByTagName('target');

@source_node_text = (); #this array will contain the text from all source nodes
@target_node_text = (); #this array will contain the text from all target nodes

	for $source_node (@source_node_list){
	$s = $source_node->textContent(); 
	# theoretically this should not be necessary, but without this step, the script throws an error 
	# textContent method not found for the object
	push (@source_node_text, $s);
	}

	for $target_node (@target_node_list){
	$t = $target_node->textContent();
	push (@target_node_text, $t);
	}

	for $source_node_text (@source_node_text){
	$source_node_text =~ tr/\x{00B6}\r\n/ /; #this removes the pilcrows and the cr/lf sequence
	$source_node_text =~ s/ {1,}/ /g; #this removes duplicate spaces
	push (@source, $source_node_text);
	}

	for $target_node_text (@target_node_text){
	$target_node_text =~ tr/\x{00B6}\r\n/ /; #this removes the pilcrows and the cr/lf sequence
	$target_node_text =~ s/ {1,}/ /g; #this removes duplicate spaces
	push (@target, $target_node_text);
	}

@source_node_list = ();
@target_node_list = ();
@source_node_text = (); 
@target_node_text = (); 

}

$outputfile = $pwd . "/" . "bi.txt"; # the output file
open (PICA, '>:encoding(utf-8)', $outputfile) || die "couldn't open the file!";

$picus = scalar(@source) - 1;
$chujus = scalar(@target) - 1;
print "Source segments: " . scalar(@source) . "\n" . "Target segments: " . scalar(@target) . "\n"; # Print all the strings

if (scalar(@source) == scalar(@target)){
print "Everything looks OK" . "\n";
}
else{
print "The number of source and target segments doesn't match" . "\n";
}

$x = 0;
while ($x <= scalar(@source))
{
print PICA @source[$x] . "\t" . @target[$x] . "\n";
$x++;
}

close PICA;