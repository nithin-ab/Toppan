import xml.etree.ElementTree as ET
from xml.dom.minidom import parseString
from datetime import datetime
import re

file=open("orderform.txt","r")
lines=file.readlines()  
file.close()

customername = '' 
masksupplier = ''
daten=''
devicen=''
siteof=''
orderformnumber = ''
revisionn = ''
pagen = ''
technologyname = ''
statusn= ''
masksetname  =''
fabunit= ''
emailaddress= ''
ponumbers=''
sitetosendmasksto=''
sitetosendinvoiceto=''
technicalcontact=''
shippingmethod=''
additionalinfo=''

for line in lines:
    
    customernames = re.search(r'\|(.*?)\s*MASK SUPPLIER', line)
    if customernames:
        customername = customernames.group(1).strip()
    
    #masksupplier_S = re.search(r'MASK SUPPLIER : (\w+)', line)
    #masksupplier_s = re.search(r'MASK SUPPLIER\s*:\s*(\w+)', line)

    masksupplier_s = re.search(r'MASK SUPPLIER\s*:\s*(.*?) DATE', line)
    if masksupplier_s:
        masksupplier = masksupplier_s.group(1).strip()

    daten_s = re.search(r'DATE\s*:\s*(.*?)\s*\|', line)
    if daten_s:
        date1 = daten_s.group(1).strip()
        d1=datetime.strptime(date1,"%m/%d/%Y")
        daten=d1.strftime("%Y-%m-%d")

    devicen_s = re.search(r'DEVICE\s*:\s*(.*?)\s*\|', line)
    if devicen_s:
        devicen = devicen_s.group(1).strip()
   
    siteof_s = re.search(r'SITE OF\s*:\s*(.*?) ORDERFORM', line)
    if siteof_s:
        siteof = siteof_s.group(1).strip()

    orderformnumber_s = re.search(r'ORDERFORM NUMBER\s*:\s*(.*?) REVISION', line)
    if orderformnumber_s:
        orderformnumber = orderformnumber_s.group(1).strip()

    revisionn_s = re.search(r'REVISION\s*:\s*(.*?) PAGE', line)
    if revisionn_s:
        revisionn = revisionn_s.group(1).strip()

    pagen_s = re.search(r'PAGE\s*:\s*(.*?) OF', line)
    if pagen_s:
        pagen = pagen_s.group(1).strip() 

    technologyname_s = re.search(r'TECHNOLOGY NAME\s*:\s*(.*?) STATUS', line)
    if technologyname_s:
        technologyname = technologyname_s.group(1).strip()

    statusn_s = re.search(r'STATUS\s*:\s*(.*?)\s*\|', line)
    if statusn_s:
        statusn = statusn_s.group(1).strip() 

    masksetname_s = re.search(r'MASK SET NAME\s*:\s*(.*?)\s*\|', line)
    if masksetname_s:
        masksetname = masksetname_s.group(1).strip()

    fabunit_s = re.search(r'FAB UNIT\s*:\s*(.*?) EMAIL', line)
    if fabunit_s:
        fabunit = fabunit_s.group(1).strip() 

    emailaddress_s = re.search(r'EMAIL\s*:\s*(.*?)\s*\|', line)
    if emailaddress_s:
        emailaddress = emailaddress_s.group(1).strip()

    ponumbers_s = re.search(r'P.O. NUMBERS\s*:\s*(.*?) SPECIFICATIONS', line)
    if ponumbers_s:
        ponumbers = ponumbers_s.group(1).strip() 

    sitetosendmasksto_s = re.search(r'SITE TO SEND MASKS TO\s*:\s*(.*?) TO THE ATTENTION OF', line)
    if sitetosendmasksto_s:
        sitetosendmasksto = sitetosendmasksto_s.group(1).strip()

    sitetosendinvoiceto_s = re.search(r'SITE TO SEND INVOICE TO\s*:\s*(.*?) TECHNICAL CONTACT', line)
    if sitetosendinvoiceto_s:
        sitetosendinvoiceto = sitetosendinvoiceto_s.group(1).strip() 

    technicalcontact_s = re.search(r'TECHNICAL CONTACT\s*:\s*(.*?) SHIPPING METHOD', line)
    if technicalcontact_s:
        technicalcontact = technicalcontact_s.group(1).strip()

    shippingmethod_s = re.search(r'SHIPPING METHOD\s*:\s*(.*?)\s*\|', line)
    if shippingmethod_s:
        shippingmethod = shippingmethod_s.group(1).strip()

    #Assuming ADDITIONAL INFORMATION as CUSTOM CLEARANCE DONE BY   
    additionalinfo_s = re.search(r'ADDITIONAL INFORMATION\s*:\s*(.*?)\s*\|', line)
    if additionalinfo_s:
        additionalinfo = additionalinfo_s.group(1).strip()


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

# Revison2
revision2=[]
for l in range(len(lines)):
    
    if "LEVEL INFORMATION" in lines[l]:   
        c1=l      

for line in lines[c1+6:]:
    if line.startswith("|XX"):
        break

    # (?:[^|]*\|) non-capturing group
    # ([^|]*) capturing group
    match1 = re.search(r'^(?:[^|]*\|){2}([^|]*)\|', line)  
    if match1:
        
        revision2.append(match1.group(1).strip())
               
# MASK CODIFICATION
numn=[] 
maskcodificationn=[] 
groupn=[]
cyclen=[]
qtyn=[]
shipdaten=[]

for l in range(len(lines)):
    
    if "MASK CODIFICATION" in lines[l]:   
        c2=l 

count1=0     
for line in lines[c2+2:]:

    if line.startswith("|----"):
        break
    
    match2 = re.search(r'^(?:[^|]*\|){1}([^|]*)\|', line)   #^ important
    if match2:
        numn.append(match2.group(1).strip())

    match3 = re.search(r'^(?:[^|]*\|){2}([^|]*)\|', line)
    if match3:
        maskcodificationn.append(match3.group(1).strip())

    match4 = re.search(r'^(?:[^|]*\|){3}([^|]*)\|', line)
    if match4:
        groupn.append(match4.group(1).strip())

    match5 = re.search(r'^(?:[^|]*\|){4}([^|]*)\|', line)
    if match5:
        cyclen.append(match5.group(1).strip())

    match6 = re.search(r'^(?:[^|]*\|){5}([^|]*)\|', line)
    if match6:
        qtyn.append(match6.group(1).strip())
    
    match7 = re.search(r'^(?:[^|]*\|){6}([^|]*)\|', line)
    if match7:
        shipdaten1=match7.group(1).strip()
        d1=datetime.strptime(shipdaten1,"%d%b%y")
        shipdaten.append(d1.strftime("%Y-%m-%d"))
        
    count1+=1     

# CRITICAL DIMENSIONS' INFORMATION
cdnumn=[]
cdnamen=[]
featuren=[]
tonen=[]
polarityn=[]

for l in range(len(lines)):
    
    if "CRITICAL DIMENSIONS' INFORMATION" in lines[l]:   
        c3=l 

count2=0     

for line in lines[c3+6:]:
    if line.startswith("|XX"):
        break

    match8 = re.search(r'^(?:[^|]*\|){7}([^|]*)\|', line)
    if match8:
        cdnumn.append(match8.group(1).strip())
    
    match9 = re.search(r'^(?:[^|]*\|){8}([^|]*)\|', line)
    if match9:
        cdnamen.append(match9.group(1).strip())
    
    match10 = re.search(r'^(?:[^|]*\|){9}([^|]*)\|', line)
    if match10:
        featuren.append(match10.group(1).strip())
    
    match11 = re.search(r'^(?:[^|]*\|){10}([^|]*)\|', line)
    if match11:
        tonen.append(match11.group(1).strip())
    
    match12 = re.search(r'^(?:[^|]*\|){11}([^|]*)\|', line)
    if match12:
        polarityn.append(match12.group(1).strip())


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
    
print("PONumbers="+ponumbers)
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


# Adding levels
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

# Adding Cdinfo
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
po_numbers.text = ponumbers

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


xml1 = ET.tostring(order_form, encoding="unicode")

x=parseString(xml1)
x1=x.toprettyxml()
print (x1)


with open("Output_python.xml", "w") as f:
    f.write(x1)

print("end")


