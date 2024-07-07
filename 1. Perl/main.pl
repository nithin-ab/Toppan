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
    my ($s, $n) = @_;

    foreach my $line (@lines) 
    {
        my $i = index($line, $s);
        
        if ($i != -1) 
        {
            my $s1 = $i-1 + length($s);
            my $result = substr($line, $s1, $n);
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

my $ponumbers=search ("P.O. NUMBERS : ",30);
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

#Revison2

for (my $l = 0; $l < scalar @lines; $l++) 
{
    if ($lines[$l] =~"LEVEL INFORMATION") 
    {
        $c1 = $l;
        #last;
    }
}

print "last line no.= $#lines\n";

for my $line (@lines[$c1+6 .. $#lines]) 
{
    last if $line =~ /^\|XX/;
    
    my $r = substr($line, 7, 3);
    $r =~ s/^\s+|\s+$//g; 
    push @revision2, $r;
    
}

print "Revision2 array= @revision2\n";

# MASK CODIFICATION

my (@numn, @maskcodificationn, @groupn, @cyclen, @qtyn, @shipdaten);

my $c2 = 0;

for (my $l = 0; $l < scalar @lines; $l++) 
{
    if ($lines[$l] =~ "MASK CODIFICATION") 
    {
        $c2 = $l;
        #last;
    }
}

my $count1 = 0;  

for my $line (@lines[$c2+2 .. $#lines]) 
{
    last if $line =~ /^\|----/;
    
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

print "num: @numn\n";
print "maskcodification: @maskcodificationn\n";
print "group: @groupn\n";
print "cycle: @cyclen\n";
print "qty: @qtyn\n";
print "shipdate: @shipdaten\n";

# CRITICAL DIMENSIONS' INFORMATION

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

my $count2 = 0; 


for my $line (@lines[$c3+6 .. $#lines]) 
{
    last if $line =~ /^\|XX/;
    
    push @cdnumn, substr($line, 37, 5) =~ s/^\s+|\s+$//gr;
    push @cdnamen, substr($line, 43, 7) =~ s/^\s+|\s+$//gr;
    push @featuren, substr($line, 51, 7) =~ s/^\s+|\s+$//gr;
    push @tonen, substr($line, 59, 7) =~ s/^\s+|\s+$//gr;
    push @polarityn, substr($line, 67, 6) =~ s/^\s+|\s+$//gr;

    $count2++;
}

print "cdnum: @cdnumn\n";
print "cdname: @cdnamen\n";
print "feature: @featuren\n";
print "tone: @tonen\n";
print "polarity: @polarityn\n";

print "PONumbers=$ponumbers\n";
print "Site to send masks to=$sitetosendmasksto\n";
print "Site to send invoice to=$sitetosendinvoiceto\n";
print "Technical contact=$technicalcontact\n";
print "Shipping method=$shippingmethod\n";
print "Additional info=$additionalinfo\n";

print "Count1=$count1\n";
print "Count2=$count2\n";


my $doc = XML::LibXML::Document->new('1.0', 'UTF-8');
my $order_form = $doc->createElement('OrderForm');
$doc->setDocumentElement($order_form);

# children
my $customer = $order_form->addNewChild('', 'Customer');
$customer->appendTextNode($customername);

my $device = $order_form->addNewChild('', 'Device');
$device->appendTextNode($devicen);

my $mask_supplier = $order_form->addNewChild('', 'MaskSupplier');
$mask_supplier->appendTextNode($masksupplier);

my $date = $order_form->addNewChild('', 'Date');
$date->appendTextNode($daten);

my $site_of = $order_form->addNewChild('', 'SiteOf');
$site_of->appendTextNode($siteof);

my $order_form_number = $order_form->addNewChild('', 'OrderFormNumber');
$order_form_number->appendTextNode($orderformnumber);

my $revision = $order_form->addNewChild('', 'Revision');
$revision->appendTextNode($revisionn);

my $page = $order_form->addNewChild('', 'Page');
$page->appendTextNode($pagen);

my $technology_name = $order_form->addNewChild('', 'TechnologyName');
$technology_name->appendTextNode($technologyname);

my $status = $order_form->addNewChild('', 'Status');
$status->appendTextNode($statusn);

my $mask_set_name = $order_form->addNewChild('', 'MaskSetName');
$mask_set_name->appendTextNode($masksetname);

my $fab_unit = $order_form->addNewChild('', 'FabUnit');
$fab_unit->appendTextNode($fabunit);

my $email_address = $order_form->addNewChild('', 'EmailAddress');
$email_address->appendTextNode($emailaddress);

# Adding levels
my $levels = $order_form->addNewChild('', 'Levels');

for (my $j = 0; $j < $count1; $j++) 
{
    my $level = $levels->addNewChild('', 'Level');
    $level->setAttribute('num', $numn[$j]);

    my $mask_codification = $level->addNewChild('', 'MaskCodification');
    $mask_codification->appendTextNode($maskcodificationn[$j]);

    my $group = $level->addNewChild('', 'Group');
    $group->appendTextNode($groupn[$j]);

    my $cycle = $level->addNewChild('', 'Cycle');
    $cycle->appendTextNode($cyclen[$j]);

    my $quantity = $level->addNewChild('', 'Quantity');
    $quantity->appendTextNode($qtyn[$j]);

    my $ship_date = $level->addNewChild('', 'ShipDate');
    $ship_date->appendTextNode($shipdaten[$j]);
}

# Adding Cdinfo
my $cdinformation = $order_form->addNewChild('', 'Cdinformation');

for (my $k = 0; $k < $count2; $k++) 
{
    my $cd_level = $cdinformation->addNewChild('', 'Level');

    my $cd_revision = $cd_level->addNewChild('', 'Revision');
    $cd_revision->appendTextNode($revision2[$k]);

    my $cd_number = $cd_level->addNewChild('', 'CDNumber');
    $cd_number->appendTextNode($cdnumn[$k]);

    my $cd_name = $cd_level->addNewChild('', 'CDName');
    $cd_name->appendTextNode($cdnamen[$k]);

    my $cd_feature = $cd_level->addNewChild('', 'Feature');
    $cd_feature->appendTextNode($featuren[$k]);

    my $cd_tone = $cd_level->addNewChild('', 'Tone');
    $cd_tone->appendTextNode($tonen[$k]);

    my $cd_polarity = $cd_level->addNewChild('', 'Polarity');
    $cd_polarity->appendTextNode($polarityn[$k]);
}

my $po_numbers = $order_form->addNewChild('', 'PONumbers');
$po_numbers->appendTextNode($ponumbers);

my $site_to_send_masks_to = $order_form->addNewChild('', 'SiteToSendMasksTo');
$site_to_send_masks_to->appendTextNode($sitetosendmasksto);

my $site_to_send_invoice_to = $order_form->addNewChild('', 'SiteToSendInvoiceTo');
$site_to_send_invoice_to->appendTextNode($sitetosendinvoiceto);

my $technical_contact = $order_form->addNewChild('', 'TechnicalContact');
$technical_contact->appendTextNode($technicalcontact);

my $shipping_method = $order_form->addNewChild('', 'ShippingMethod');
$shipping_method->appendTextNode($shippingmethod);

my $additional_information = $order_form->addNewChild('', 'AdditionalInformation');
$additional_information->appendTextNode($additionalinfo);

# String conversion
my $xmlstr = $doc->toString(1);
print $xmlstr;

open(my $fh1, '>', 'Output_perl.xml') or die "Could not open file 'Output_perl.xml' $!";
print $fh1 $xmlstr;
close $fh1;

print "End\n";
