Attribute VB_Name = "Module2"
Sub main()

   Dim filepath As String
   Dim file As String
   Dim lines() As String
   Dim fileno As Integer
   Dim regex As Object

   Dim masksupplier As String
   Dim customername As String
   Dim daten1 As String
   Dim daten2() As String
   Dim daten As String
   Dim devicen As String
   Dim siteof As String
   Dim orderformnumber As String
   Dim revisionn As String
   Dim pagen As String
   Dim technologyname As String
   Dim statusn As String
   Dim masksetname As String
   Dim fabunit As String
   Dim emailaddress As String
   Dim ponumbers As String
   Dim sitetosendmasksto As String
   Dim sitetosendinvoiceto As String
   Dim technicalcontact As String
   Dim shippingmethod As String
   Dim additionalinfo As String
   
   Dim revision2 As New Collection

   Dim numn As New Collection
   Dim maskcodificationn As New Collection
   Dim groupn As New Collection
   Dim cyclen As New Collection
   Dim qtyn As New Collection
   Dim shipdaten As New Collection
   Dim shipdaten1 as String
   Dim shipdaten2 as String
   Dim day as String
   Dim month as String
   Dim year as String

   Dim cdnumn As New Collection
   Dim cdnamen As New Collection
   Dim featuren As New Collection
   Dim tonen As New Collection
   Dim polarityn As New Collection

   Dim count1 as Integer
   Dim count2 as Integer
  
   filepath = "C:\Users\nithi\OneDrive\Desktop\Toppan\orderform.txt"
   fileno = FreeFile

   Open filepath For Input As #fileno
   file = Input$(LOF(fileno), fileno)
   Close #fileno
   lines = Split(file, vbCrLf)

   Dim i As Integer
   Dim line As String

   Set regex = CreateObject("VBScript.RegExp")
   regex.Global = False  ' stop after 1st match

For i = LBound(lines) To UBound(lines)
        line = lines(i)


      regex.Pattern = "\|(.*?)\s*MASK SUPPLIER"
      If regex.Test(line) Then
            customername = regex.Execute(line)(0).SubMatches(0)
            customername = Trim(customername)
      End If

      regex.Pattern = "MASK SUPPLIER\s*:\s*(.*?) DATE"
      If regex.Test(line) Then
            masksupplier = Trim(regex.Execute(line)(0).SubMatches(0))
      End If

      regex.Pattern = "DATE\s*:\s*(.*?)\s*\|"
      If regex.Test(line) Then
            daten1 = Trim(regex.Execute(line)(0).SubMatches(0))
            daten2=Split(daten1,"/")
            daten=daten2(2) & "-" & daten2(0) & "-" & daten2(1)

      End If

      regex.Pattern = "DEVICE\s*:\s*(.*?)\s*\|"
        If regex.test(line) Then
            devicen = regex.Execute(line)(0).submatches(0)
            devicen = Trim(devicen)
        End If
   
      regex.Pattern = "SITE OF\s*:\s*(.*?) ORDERFORM"
      If regex.Test(line) Then
         siteof = Trim(regex.Execute(line)(0).SubMatches(0))
      End If

      regex.Pattern = "ORDERFORM NUMBER\s*:\s*(.*?) REVISION"
      If regex.Test(line) Then
            orderformnumber = Trim(regex.Execute(line)(0).SubMatches(0))
      End If

      regex.Pattern = "REVISION\s*:\s*(.*?) PAGE"
      If regex.Test(line) Then
            revisionn = Trim(regex.Execute(line)(0).SubMatches(0))
      End If

      regex.Pattern = "PAGE\s*:\s*(.*?) OF"
      If regex.Test(line) Then
            pagen = Trim(regex.Execute(line)(0).SubMatches(0))
      End If

      regex.Pattern = "TECHNOLOGY NAME\s*:\s*(.*?) STATUS"
        If regex.Test(line) Then
            technologyname = Trim(regex.Execute(line)(0).SubMatches(0))
        End If

      regex.Pattern = "STATUS\s*:\s*(.*?)\s*\|"
        If regex.Test(line) Then
            statusn = Trim(regex.Execute(line)(0).SubMatches(0))
        End If

      regex.Pattern = "MASK SET NAME\s*:\s*(.*?)\s*\|"
      If regex.Test(line) Then
            masksetname = Trim(regex.Execute(line)(0).SubMatches(0))
      End If

      regex.Pattern = "FAB UNIT\s*:\s*(.*?) EMAIL"
      If regex.Test(line) Then
            fabunit = Trim(regex.Execute(line)(0).SubMatches(0))
      End If
   
      regex.Pattern = "EMAIL\s*:\s*(.*?)\s*\|"
      If regex.Test(line) Then
            emailaddress = Trim(regex.Execute(line)(0).SubMatches(0))
      End If

      regex.Pattern = "P.O. NUMBERS\s*:\s*(.*?) SPECIFICATIONS"
      If regex.Test(line) Then
            ponumbers = Trim(regex.Execute(line)(0).SubMatches(0))
            
      End If

      regex.Pattern = "SITE TO SEND MASKS TO\s*:\s*(.*?) TO THE ATTENTION OF"
      If regex.Test(line) Then
            sitetosendmasksto = Trim(regex.Execute(line)(0).SubMatches(0))
      End If

      regex.Pattern = "SITE TO SEND INVOICE TO\s*:\s*(.*?) TECHNICAL CONTACT"
      If regex.Test(line) Then
            sitetosendinvoiceto = Trim(regex.Execute(line)(0).SubMatches(0))
      End If

      regex.Pattern = "TECHNICAL CONTACT\s*:\s*(.*?) SHIPPING METHOD"
      If regex.Test(line) Then
            technicalcontact = Trim(regex.Execute(line)(0).SubMatches(0))
      End If

      regex.Pattern = "SHIPPING METHOD\s*:\s*(.*?)\s*\|"
      If regex.Test(line) Then
            shippingmethod = Trim(regex.Execute(line)(0).SubMatches(0))
      End If
      regex.Pattern = "ADDITIONAL INFORMATION\s*:\s*(.*?)\s*\|"
      If regex.Test(line) Then
         additionalinfo = Trim(regex.Execute(line)(0).SubMatches(0))
      End If
      

      'Revision2
      If InStr(line, "LEVEL INFORMATION") > 0 Then
            'Debug.print InStr(line, "LEVEL INFORMATION") 
      
         Dim c1 As Integer
         c1 = i

         For c1 = i+6 To UBound(lines)
            line = lines(c1)
            If Left(line, 3) = "|XX" Then Exit For
               regex.Pattern = "^(?:[^|]*\|){2}([^|]*)\|"
               If regex.Test(line) Then
                  revision2.add Trim(regex.Execute(line)(0).SubMatches(0))
               End If
         Next c1
         
      End If

     'MASK CODIFICATION

      If InStr(line, "MASK CODIFICATION") > 0 Then
                  
         Dim c2 As Integer
         c2 = i
         count1=0

         For c2 = i+2 To UBound(lines)
            line = lines(c2)
            If Left(line, 5) = "|----" Then Exit For

            regex.Pattern = "^(?:[^|]*\|){1}([^|]*)\|"
            If regex.Test(line) Then
               numn.add Trim(regex.Execute(line)(0).SubMatches(0))
            End If

            regex.Pattern = "^(?:[^|]*\|){2}([^|]*)\|"
            If regex.Test(line) Then
                  maskcodificationn.add Trim(regex.Execute(line)(0).SubMatches(0))
            End If

            regex.Pattern = "^(?:[^|]*\|){3}([^|]*)\|"
            If regex.Test(line) Then
                  groupn.add Trim(regex.Execute(line)(0).SubMatches(0))
            End If

            regex.Pattern = "^(?:[^|]*\|){4}([^|]*)\|"
            If regex.Test(line) Then
                  cyclen.add Trim(regex.Execute(line)(0).SubMatches(0))
            End If

            regex.Pattern = "^(?:[^|]*\|){5}([^|]*)\|"
            If regex.Test(line) Then
                  qtyn.add Trim(regex.Execute(line)(0).SubMatches(0))
            End If

            regex.Pattern = "^(?:[^|]*\|){6}([^|]*)\|"
            If regex.Test(line) Then
                  shipdaten1=Trim(regex.Execute(line)(0).SubMatches(0))
                  day = Left(shipdaten1, 2)
                  month = Mid(shipdaten1, 3, 3)
                  year = Right(shipdaten1, 2)
                  year = "20" &year

                  Select Case month 
                     Case "JAN": month = "01"
                     Case "FEB": month = "02"
                     Case "MAR": month = "03"
                     Case "APR": month = "04"
                     Case "MAY": month = "05"
                     Case "JUN": month = "06"
                     Case "JUL": month = "07"
                     Case "AUG": month = "08"
                     Case "SEP": month = "09"
                     Case "OCT": month = "10"
                     Case "NOV": month = "11"
                     Case "DEC": month = "12"
                     'Case Else :
                  End Select

                  shipdaten2= year & "-" & month & "-" & day
                  shipdaten.add shipdaten2

            End If
            count1=count1+1

         Next c2 
      
      End If
Next i 

For i = LBound(lines) To UBound(lines)
      line = lines(i)
      'CRITICAL DIMENSIONS' INFORMATION
      
      If InStr(line, "CRITICAL DIMENSIONS' INFORMATION") > 0 Then
      
         Dim c3 As Integer
         c3 = i
         count2=0
         For c3 = i+6 To UBound(lines)
            line = lines(c3)
            If Left(line, 3) = "|XX" Then Exit For
               
               regex.Pattern = "^(?:[^|]*\|){7}([^|]*)\|"
               If regex.Test(line) Then
                  cdnumn.add Trim(regex.Execute(line)(0).SubMatches(0))
               End If

               regex.Pattern = "^(?:[^|]*\|){8}([^|]*)\|"
               If regex.Test(line) Then
                  cdnamen.add Trim(regex.Execute(line)(0).SubMatches(0))
               End If

               regex.Pattern = "^(?:[^|]*\|){9}([^|]*)\|"
               If regex.Test(line) Then
                  featuren.add Trim(regex.Execute(line)(0).SubMatches(0))
               End If

               regex.Pattern = "^(?:[^|]*\|){10}([^|]*)\|"
               If regex.Test(line) Then
                  tonen.add Trim(regex.Execute(line)(0).SubMatches(0))
               End If

               regex.Pattern = "^(?:[^|]*\|){11}([^|]*)\|"
               If regex.Test(line) Then
                  polarityn.add Trim(regex.Execute(line)(0).SubMatches(0))
               End If

         count2=count2+1

         Next c3
         
      End If

Next i

Debug.Print "Customer name:" & customername
Debug.Print "Mask supplier:" & masksupplier
Debug.Print "Date:" & daten
Debug.Print "Device:" & devicen
Debug.Print "Site of:" & siteof
Debug.Print "Orderform no.:" & orderformnumber
Debug.Print "Revision:" & revisionn
Debug.Print "Page:" & pagen
Debug.Print "Technology name:" & technologyname
Debug.Print "Status:" & statusn
Debug.Print "Mask set name:" & masksetname
Debug.Print "Fab unit:" & fabunit
Debug.Print "Email address:" & emailaddress
Debug.Print "PO numbers:" & ponumbers
Debug.Print "Site to send masks to:" & sitetosendmasksto
Debug.Print "Site to send invoive to:" & sitetosendinvoiceto
Debug.Print "Technical contact:" & technicalcontact
Debug.Print "Shipping method:" & shippingmethod
Debug.Print "Additional info:" & additionalinfo


For Each k In revision2
   Debug.Print "revision2:" & k
   Next k
For Each k In numn
   Debug.Print "numn:" & k
   Next k

For Each k In maskcodificationn
   Debug.Print "mask codification:" & k
   Next k
For Each k In groupn
   Debug.Print "GRP:" & k
   Next k
For Each k In cyclen
   Debug.Print "CYL:" & k
   Next k
For Each k In qtyn
   Debug.Print "QTY:" & k
   Next k
For Each k In shipdaten
   Debug.Print "SHIP DATE:" & k
   Next k

For Each k In cdnumn
   Debug.Print "CD NUM:" & k
   Next k

For Each k In cdnamen
   Debug.Print "CD NAME:" & k
   Next k

For Each k In featuren
   Debug.Print "FEATURE:" & k
   Next k
For Each k In tonen
   Debug.Print "TONE:" & k
   Next k
For Each k In polarityn
   Debug.Print "POLARITY:" & k
   Next k

Debug.Print "Count1:" & count1
Debug.Print "Count2:" & count2


Dim x As Object
Dim root As Object
Dim customer As Object
Dim device As Object
Dim mask_supplier As Object
Dim dateobj As Object
Dim site_of As Object
Dim order_form_number As Object
Dim revision As Object
Dim page As Object
Dim technology_name As Object
Dim status As Object
Dim mask_set_name As Object
Dim fab_unit As Object
Dim email_address As Object
Dim levels As Object
Dim cdinformation As Object
Dim po_numbers As Object
Dim site_to_send_masks_to As Object
Dim site_to_send_invoice_to As Object
Dim technical_contact As Object
Dim shipping_method As Object
Dim additional_information As Object
Dim level As Object
Dim cd_level As Object
    

Set x = CreateObject("MSXML2.DOMDocument")
Set root = x.createElement("OrderForm")
x.appendChild root

Set customer = x.createElement("Customer")
customer.Text = customername
root.appendChild customer

Set device = x.createElement("Device")
device.Text = devicen
root.appendChild device

Set mask_supplier = x.createElement("MaskSupplier")
mask_supplier.Text = masksupplier
root.appendChild mask_supplier

Set dateobj = x.createElement("Date")
dateobj.Text = daten
root.appendChild dateobj

Set site_of = x.createElement("SiteOf")
site_of.Text = siteof
root.appendChild site_of

Set order_form_number = x.createElement("OrderFormNumber")
order_form_number.Text = orderformnumber
root.appendChild order_form_number

Set revision = x.createElement("Revision")
revision.Text = revisionn
root.appendChild revision

Set page = x.createElement("Page")
page.Text = pagen
root.appendChild page

Set technology_name = x.createElement("TechnologyName")
technology_name.Text = technologyname
root.appendChild technology_name

Set status = x.createElement("Status")
status.Text = statusn
root.appendChild status

Set mask_set_name = x.createElement("MaskSetName")
mask_set_name.Text = masksetname
root.appendChild mask_set_name

Set fab_unit = x.createElement("FabUnit")
fab_unit.Text = fabunit
root.appendChild fab_unit

Set email_address = x.createElement("EmailAddress")
email_address.Text = emailaddress
root.appendChild email_address


'levels
Set levels = x.createElement("Levels")
root.appendChild levels


For j = 1 To count1    '(imp)
   Set level = x.createElement("Level")
   level.setAttribute "num", numn(j)
   
   Dim mask_codification As Object
   Set mask_codification = x.createElement("MaskCodification")
   mask_codification.Text = maskcodificationn(j)
   level.appendChild mask_codification
   
   Dim group As Object
   Set group = x.createElement("Group")
   group.Text = groupn(j)
   level.appendChild group
   
   Dim cycle As Object
   Set cycle = x.createElement("Cycle")
   cycle.Text = cyclen(j)
   level.appendChild cycle
   
   Dim quantity As Object
   Set quantity = x.createElement("Quantity")
   quantity.Text = qtyn(j)
   level.appendChild quantity
   
   Dim ship_date As Object
   Set ship_date = x.createElement("ShipDate")
   ship_date.Text = shipdaten(j)
   level.appendChild ship_date
   
   levels.appendChild level

Next j

'Cdinfo

Set cdinformation = x.createElement("Cdinformation")
root.appendChild cdinformation

For k = 1 To count2

   Set cd_level = x.createElement("Level")
   
   Dim cdrevision As Object
   Set cdrevision = x.createElement("Revision")
   cdrevision.Text = revision2(k)
   cd_level.appendChild cdrevision
   
   Dim cdnumber As Object
   Set cdnumber = x.createElement("CDNumber")
   cdnumber.Text = cdnumn(k)
   cd_level.appendChild cdnumber
   
   Dim cdname As Object
   Set cdname = x.createElement("CDName")
   cdname.Text = cdnamen(k)
   cd_level.appendChild cdname
   
   Dim cdfeature As Object
   Set cdfeature = x.createElement("Feature")
   cdfeature.Text = featuren(k)
   cd_level.appendChild cdfeature
   
   Dim cdtone As Object
   Set cdtone = x.createElement("Tone")
   cdtone.Text = tonen(k)
   cd_level.appendChild cdtone
   
   Dim cdpolarity As Object
   Set cdpolarity = x.createElement("Polarity")
   cdpolarity.Text = polarityn(k)
   cd_level.appendChild cdpolarity
   
   cdinformation.appendChild cd_level
   
Next k

Set po_numbers = x.createElement("PONumbers")
po_numbers.Text = ponumbers
root.appendChild po_numbers

Set site_to_send_masks_to = x.createElement("SiteToSendMasksTo")
site_to_send_masks_to.Text = sitetosendmasksto
root.appendChild site_to_send_masks_to

Set site_to_send_invoice_to = x.createElement("SiteToSendInvoiceTo")
site_to_send_invoice_to.Text = sitetosendinvoiceto
root.appendChild site_to_send_invoice_to

Set technical_contact = x.createElement("TechnicalContact")
technical_contact.Text = technicalcontact
root.appendChild technical_contact

Set shipping_method = x.createElement("ShippingMethod")
shipping_method.Text = shippingmethod
root.appendChild shipping_method

Set additional_information = x.createElement("AdditionalInformation")
additional_information.Text = additionalinfo
root.appendChild additional_information


filePath = ThisWorkbook.Path & "\Output_vba.xml"
x.Save filePath

Debug.print("End")

   
End sub