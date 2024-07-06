use strict;
use warnings;
use DateTime::Format::Strptime;

use XML::LibXML;

my $file = 'orderform.txt';

open(my $fh, '<', $file) or die "Could not open file '$file";

my @lines = <$fh>;
#print"@lines";

close($fh);

sub search 
{
    my ($search_string, $n) = @_;

    foreach my $line (@lines) 
    {
        my $i = index($line, $search_string);
        
        if ($i != -1) 
        {
            my $result_start_index = $i + length($search_string);
            my $result = substr($line, $result_start_index, $n);
            $result=~ s/^\s+|\s+$//g;
            return $result; 
        }
    }

    return "";
}

my $customername = search("!",30);
my $devicen = search("DEVICE : ",30);
my $masksupplier = search("MASK SUPPLIER : ", 30);

my $daten1 = search ("DATE : " ,11);
print ("daten1= $daten1\n");
my $daten2 = DateTime::Format::Strptime->new(pattern => '%m/%d/%Y',on_error=>"croak");
print"daten2=$daten2\n";
my $daten3 = $daten2->parse_datetime($daten1);
print"daten3=$daten3\n";
my $daten = $daten3->strftime('%Y-%m-%d');

my $siteof= search ("SITE OF : ",30);
my $orderformnumber = search("ORDERFORM NUMBER : ",15);
my $revisionn = search("REVISION : ",10);
my $pagen = search ("PAGE : ",4);
my $technologyname = search ("TECHNOLOGY NAME : ",20);
my $statusn= search("STATUS    : ",5);
my $masksetname  =search ("MASK SET NAME : ",30);
my $fabunit= search("FAB UNIT        : ",20);
my $emailaddress= search("EMAIL : ",22);

my $PONumbers=search ("P.O. NUMBERS : ",30);
my $sitetosendmasksto=search("SITE TO SEND MASKS TO : ",24);
my $sitetosendinvoiceto=search("SITE TO SEND INVOICE TO : ",22);
my $technicalcontact=search("TECHNICAL CONTACT : ",26);
my $shippingmethod=search("SHIPPING METHOD : ",15);
my $additionalinfo=search("ADDITIONAL INFORMATION : ",15);

print("Customer= $customername\n");
print("Device= $devicen\n");
print("Mask supplier= $masksupplier\n");
print("Date= $daten\n");
print("Site of= $siteof\n");
print("Order form number= $orderformnumber\n");
print("Revision= $revisionn\n");
print("Page= $pagen\n");
print("Technolgy name= $technologyname\n");
print("Status= $statusn\n");
print("Mask set name= $masksetname\n");
print("Fab Unit= $fabunit\n");
print("Email address= $emailaddress\n");


my @revision2;

my $c1 = 0;

#Code for Revison2
for (my $l = 0; $l < scalar @lines; $l++) 
{
    if ($lines[$l] =~"LEVEL INFORMATION") 
    {
        $c1 = $l;
        #last;
    }
}

my $t1 = 0;
print "$#lines\n";

for my $line (@lines[$c1 .. $#lines]) 
{
    last if $line =~ /^\|XX/;
    if ($t1 >= 6) 
    {
        my $r = substr($line, 7, 3);
        $r =~ s/^\s+|\s+$//g; 
        push @revision2, $r;
    }
    $t1++;
}

print "Revision2: @revision2\n";

#Code for elements inside MASK CODIFICATION


my (@numn, @maskcodificationn, @groupn, @cyclen, @qtyn, @shipdaten);

my $c2 = 0;

# Find the line containing "MASK CODIFICATION"
for (my $l = 0; $l < scalar @lines; $l++) 
{
    if ($lines[$l] =~ "MASK CODIFICATION") 
    {
        $c2 = $l;
        #last;
    }
}

my $t2 = 0;
my $count1 = 0;  # Counter for knowing no. of lines in maskcodification

# Create a DateTime parser for ship date

for my $line (@lines[$c2 + 2 .. $#lines]) 
{
    last if $line =~ /^\|----/;
    if ($t2 >= 0) 
    {
        push @numn, substr($line, 3, 4) =~ s/^\s+|\s+$//gr;
        push @maskcodificationn, substr($line, 8, 33) =~ s/^\s+|\s+$//gr;
        push @groupn, substr($line, 42, 3) =~ s/^\s+|\s+$//gr;
        push @cyclen, substr($line, 47, 3) =~ s/^\s+|\s+$//gr;
        push @qtyn, substr($line, 52, 3) =~ s/^\s+|\s+$//gr;

        my $shipdaten1 = substr($line, 56, 8) =~ s/^\s+|\s+$//gr;
        print "shipdaten1= $shipdaten1\n";
        my $shipdaten2 = DateTime::Format::Strptime->new(pattern => '%d%b%y',on_error=>"croak");
        print "shipdaten2= $shipdaten2\n";
        my $shipdaten3 = $shipdaten2->parse_datetime($shipdaten1);
        print "shipdaten2= $shipdaten2\n";
        my $shipdaten = $shipdaten3->strftime('%Y-%m-%d');

        push @shipdaten, $shipdaten;


        $count1++;
    }
    $t2++;
}


print "num: @numn\n";
print "maskcodification: @maskcodificationn\n";
print "group: @groupn\n";
print "cycle: @cyclen\n";
print "qty: @qtyn\n";
print "shipdate: @shipdaten\n";

#Code for elements inside CRITICAL DIMENSIONS' INFORMATION

my (@cdnumn, @cdnamen, @featuren, @tonen, @polarityn);

my $c3 = 0;

for (my $l = 0; $l < scalar @lines; $l++) 
{
    if ($lines[$l] =~ "CRITICAL DIMENSIONS' INFORMATION") 
    {
        $c3 = $l;
        #last;
    }
}

my $t3 = 0;
my $count2 = 0;  # Counter for knowing no. of lines in critical dimensions' information


for my $line (@lines[$c3 .. $#lines]) 
{
    last if $line =~ /^\|XX/;
    if ($t3 >= 6) 
    {
        push @cdnumn, substr($line, 37, 5) =~ s/^\s+|\s+$//gr;
        push @cdnamen, substr($line, 43, 7) =~ s/^\s+|\s+$//gr;
        push @featuren, substr($line, 51, 7) =~ s/^\s+|\s+$//gr;
        push @tonen, substr($line, 59, 7) =~ s/^\s+|\s+$//gr;
        push @polarityn, substr($line, 67, 6) =~ s/^\s+|\s+$//gr;
        $count2++;
    }
    $t3++;
}

print "cdnum: @cdnumn\n";
print "cdname: @cdnamen\n";
print "feature: @featuren\n";
print "tone: @tonen\n";
print "polarity: @polarityn\n";

print "PONumbers=$PONumbers\n";
print "Site to send masks to=$sitetosendmasksto\n";
print "Site to send invoice to=$sitetosendinvoiceto\n";
print "Technical contact=$technicalcontact\n";
print "Shipping method=$shippingmethod\n";
print "Additional info=$additionalinfo\n";

print "Count1=$count1\n";
print "Count2=$count2\n";

