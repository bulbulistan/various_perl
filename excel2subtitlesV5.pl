use File::Glob ':glob';
use Cwd;
use XML::LibXML;
use XML::Twig;
use Win32::OLE;
use Win32::OLE qw(in with);
use Win32::OLE::Variant;
use Win32::OLE::Const 'Microsoft Excel';

$adresar = getcwd();
$outxml = $adresar . "/" . "BrightCove_subtitles.xml";

########### convert excel to txt ##########

@excelfajly = glob("*.xlsx");
my $Excel = Win32::OLE->GetActiveObject('Excel.Application') || Win32::OLE->new('Excel.Application', 'Quit');

for ( $i = 0 ; $i < scalar(@excelfajly) ; $i++ ) {

my $filename = @excelfajly[$i];
$filename =~ s/\..*/\.txt/;

my $sourcefile = $adresar . "/" . @excelfajly[$i];
my $target = $adresar . "/" . $filename;

$target =~ s/\//\\/g; #pojebany windows
print $target . "\n";

my $Book = $Excel->Workbooks->Open($sourcefile); 
$Book->SaveAs({Filename=>$target, FileFormat=>xlUnicodeText});
$Book->Close({SaveChanges=>False});
}
$Excel->Quit();


#########################################
########## convert txt to srt / xml #####

@fajly =  glob("*.txt");

$dom = XML::LibXML::Document->new( "1.0", "UTF-8" );
$root = $dom->createElement("tt");

$root->setAttribute("xmlns", "http://www.w3.org/ns/ttml");
$root->setAttribute("xmlns:tts", "http://www.w3.org/ns/ttml#styling");
$root->setAttribute("xml:lang", "");

$dom->setDocumentElement( $root );
$root = $dom->documentElement();

$head = $dom->createElement("head");
$body = $dom->createElement("body");
$styling = $dom->createElement("styling");

$style = $dom->createElement("style");
$style->setAttribute("tts:color", "white");
$style->setAttribute("tts:fontFamily", "Verdana");
$style->setAttribute("tts:fontSize", "100%");
$style->setAttribute("tts:lineHeight", "8%");
$style->setAttribute("tts:textOutline", "#333333 8% 8%");
$style->setAttribute("xml:id","1");

$styling->appendChild($style);
$head->appendChild($styling);
$root->appendChild($head);
$root->appendChild($body); #fix

for ( $i = 0 ; $i < scalar(@fajly) ; $i++ ) {

$div = $dom->createElement("div");

$lang = @fajly[$i];
#$lang =~ s/.*?_(.*?)\..*/$1/;
$lang =~ s/\..*//g;

$div->setAttribute("xml:lang", $lang);
$body->appendChild( $div );


$in = @fajly[$i];
open( KOKOT, '<:encoding(utf-16le)', $in )  || die "neni subor, chuju";

$fn = @fajly[$i];
$fn =~ s/\.txt/\.srt/;

$out = $fn;

open( PICA, '>:encoding(UTF-8)', $out ) || die "couldn't open the file!";
$x = 0;
$cl = 0;

while ($rit = <KOKOT>) {
$cl++;

if ($cl == 1) #firstline has no time code, i.e. no digits
{
print $rit . "\n";
}

else
{
$x++;
chomp($rit);

$rit =~ s/\-\-\>/\t/;
$rit =~ s/\t+/\t/g;

@riadok = split('\t', $rit); 

$start = @riadok[0];
$end = @riadok[1];
$text = @riadok[2];

$start =~ s/\"//g;
$start =~ s/\,/\./g;
$start =~ s/ //g;
$end =~ s/\"//g;
$end =~ s/\,/\./g;
$end =~ s/ //g;
$text =~ s/\"//g;
$text =~ s/\r//g; #bez toho to tam dava jeden CR navyse 

print PICA $x . "\n" . $start . " --> " . $end . "\n" . $text . "\n\n"; #na konci 2x\n lebo taky je format SRT

$p = $dom->createElement("p");
$p->setAttribute("begin", $start);
$p->setAttribute("end", $end);
$p->setAttribute("style", "1");

$cd = $dom->createCDATASection( $text );
$p->appendChild($cd);

$div->appendChild($p);

}
}
}

$state = $dom->toFile($outxml, 1);
