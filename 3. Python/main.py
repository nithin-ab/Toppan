import xml.etree.ElementTree as ET
from xml.dom.minidom import parseString
from datetime import datetime

file=open("orderform.txt","r")
lines=file.readlines()
file.close()

def search (search_string, n):

    for line in lines:
        
        i = line.find(search_string)
        if i!=-1:
            
            result_start_index = i-1 + len(search_string)
            return line[result_start_index:result_start_index+n].strip()
        
    return ""      
          
customername = search("!",30)
devicen = search("DEVICE : ",30)
masksupplier = search("MASK SUPPLIER : ", 30)

daten1 = search ("DATE : " ,11)
d1=datetime.strptime(daten1,"%m/%d/%Y")
daten=d1.strftime("%Y-%m-%d")

siteof= search ("SITE OF : ",30)
orderformnumber = search("ORDERFORM NUMBER : ",15)
revisionn = search("REVISION : ",10)
pagen = search ("PAGE : ",4)
technologyname = search ("TECHNOLOGY NAME : ",20)
statusn= search("STATUS    : ",5)
masksetname  =search ("MASK SET NAME : ",30)
fabunit= search("FAB UNIT        : ",20)
emailaddress= search("EMAIL : ",22)

PONumbers=search ("P.O. NUMBERS : ",30)
sitetosendmasksto=search("SITE TO SEND MASKS TO : ",24)
sitetosendinvoiceto=search("SITE TO SEND INVOICE TO : ",22)
technicalcontact=search("TECHNICAL CONTACT : ",26)
shippingmethod=search("SHIPPING METHOD : ",15)
additionalinfo=search("ADDITIONAL INFORMATION : ",15)

print("Customer="+customername)
print("Device="+devicen)
print("Mask supplier="+masksupplier)
print("Date="+daten)
print("Site of="+siteof)
print("Order form number="+orderformnumber)
print("Revision="+revisionn)
print("Page="+pagen)
print("Technolgy name="+technologyname)
print("Status="+statusn)
print("Mask set name="+masksetname)
print("Fab Unit="+fabunit)
print("Email address="+emailaddress)

#Code for Revison2
revision2=[]
for l in range(len(lines)):
    
    if "LEVEL INFORMATION" in lines[l]:   
        c1=l      

for line in lines[c1+6:]:
    if line.startswith("|XX"):
        break
    
    revision2.append(line[7:10].strip())
               
          
#Code for elements inside MASK CODIFICATION
numn=[] 
maskcodificationn=[] 
groupn=[]
cyclen=[]
qtyn=[]
shipdaten=[]

for l in range(len(lines)):
    
    if "MASK CODIFICATION" in lines[l]:   
        c2=l 

count1=0     #counter for knowing no.of lines in maskmodification
for line in lines[c2+2:]:

    if line.startswith("|----"):
        break
    
    numn.append(line[3:7].strip())
    maskcodificationn.append(line[8:41].strip())
    groupn.append(line[42:45].strip())
    cyclen.append(line[47:50].strip())
    qtyn.append(line[52:55].strip())

    shipdaten1=(line[56:64].strip())
    d1=datetime.strptime(shipdaten1,"%d%b%y")
    shipdaten.append(d1.strftime("%Y-%m-%d"))
        
    count1+=1     

#Code for elements inside CRITICAL DIMENSIONS' INFORMATION
cdnumn=[]
cdnamen=[]
featuren=[]
tonen=[]
polarityn=[]

for l in range(len(lines)):
    
    if "CRITICAL DIMENSIONS' INFORMATION" in lines[l]:   
        c3=l 

count2=0     #counter for knowing no.of lines in level information/critical dimensions' information

for line in lines[c3+6:]:
    if line.startswith("|XX"):
        break
    
    cdnumn.append(line[37:42].strip())
    cdnamen.append(line[43:50].strip())
    featuren.append(line[51:58].strip())
    tonen.append(line[59:66].strip())
    polarityn.append(line[67:73].strip())

    count2+=1
            
print("num=",numn)
print("maskcodification=",maskcodificationn) 
print("group=",groupn)
print("cycle=",cyclen)
print("Qty=",qtyn)
print("Shipdate=",shipdaten)
print("Revision2=",revision2)
print("cdnum=",cdnumn)
print("CD Name=",cdnamen)
print("Feature=",featuren)
print("Tone=",tonen)
print("Polarity=",polarityn)
    
print("PONumbers="+PONumbers)
print("Site to send masks to="+sitetosendmasksto)
print("Site to send invoice to="+sitetosendinvoiceto)
print("Technical contact="+technicalcontact)
print("Shipping method="+shippingmethod)
print("Additional info="+additionalinfo)

print("Count1=",count1)
print("Count2=",count2)

#Root
order_form = ET.Element("OrderForm")  

# Children
customer = ET.SubElement(order_form, "Customer")
customer.text = customername

device = ET.SubElement(order_form, "Device")
device.text = devicen

mask_supplier = ET.SubElement(order_form, "MaskSupplier")
mask_supplier.text = masksupplier

date = ET.SubElement(order_form, "Date")
date.text = daten

site_of = ET.SubElement(order_form, "SiteOf")
site_of.text = siteof

order_form_number = ET.SubElement(order_form, "OrderFormNumber")
order_form_number.text = orderformnumber

revision = ET.SubElement(order_form, "Revision")
revision.text = revisionn

page = ET.SubElement(order_form, "Page")
page.text = pagen

technology_name = ET.SubElement(order_form, "TechnologyName")
technology_name.text = technologyname

status = ET.SubElement(order_form, "Status")
status.text = statusn

mask_set_name = ET.SubElement(order_form, "MaskSetName")
mask_set_name.text = masksetname

fab_unit = ET.SubElement(order_form, "FabUnit")
fab_unit.text = fabunit

email_address = ET.SubElement(order_form, "EmailAddress")
email_address.text = emailaddress


# Adding Levels
levels = ET.SubElement(order_form, "Levels")

j=0
while(j<count1):
    
    level = ET.SubElement(levels, "Level", num=numn[j])
    
    mask_codification = ET.SubElement(level, "MaskCodification")
    mask_codification.text = maskcodificationn[j]

    group = ET.SubElement(level, "Group")
    group.text = groupn[j]

    cycle = ET.SubElement(level, "Cycle")
    cycle.text = cyclen[j]

    quantity = ET.SubElement(level, "Quantity")
    quantity.text = qtyn[j]

    ship_date = ET.SubElement(level, "ShipDate")
    ship_date.text = shipdaten[j]

    j+=1

# Adding Cdinformation
cdinformation = ET.SubElement(order_form, "Cdinformation")

k=0
while (k<count2):

    cd_level = ET.SubElement(cdinformation, "Level")

    cd_revision = ET.SubElement(cd_level, "Revision")
    cd_revision.text = revision2[k]

    cd_number = ET.SubElement(cd_level, "CDNumber")
    cd_number.text = cdnumn[k]

    cd_name = ET.SubElement(cd_level, "CDName")
    cd_name.text = cdnamen[k]

    cd_feature = ET.SubElement(cd_level, "Feature")
    cd_feature.text = featuren[k]

    cd_tone = ET.SubElement(cd_level, "Tone")
    cd_tone.text = tonen[k]

    cd_polarity = ET.SubElement(cd_level, "Polarity")
    cd_polarity.text = polarityn[k]

    k+=1

po_numbers = ET.SubElement(order_form, "PONumbers")
po_numbers.text = PONumbers

site_to_send_masks_to = ET.SubElement(order_form, "SiteToSendMasksTo")
site_to_send_masks_to.text = sitetosendmasksto

site_to_send_invoice_to = ET.SubElement(order_form, "SiteToSendInvoiceTo")
site_to_send_invoice_to.text = sitetosendinvoiceto

technical_contact = ET.SubElement(order_form, "TechnicalContact")
technical_contact.text = technicalcontact

shipping_method = ET.SubElement(order_form, "ShippingMethod")
shipping_method.text = shippingmethod

additional_information = ET.SubElement(order_form, "AdditionalInformation")
additional_information.text = additionalinfo


xmlstr = ET.tostring(order_form, encoding="unicode")

x=parseString(xmlstr)
x1=x.toprettyxml()
print (x1)


with open("Output_python.xml", "w") as f:
    f.write(x1)

print("XML file generated successfully.")

