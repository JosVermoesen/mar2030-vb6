Attribute VB_Name = "modPeppol"
Option Explicit
DefInt A-Z

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef pguid As GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (ByRef rguid As GUID, ByVal lpsz As Long, ByVal cchMax As Long) As Long

Public Function CheckforAmp(toCheck As String) As String

    If InStr(toCheck, "&") Then 'verbeteren voor XML bestand!!!
        CheckforAmp = Replace(toCheck, "&", "&amp;")
    Else
        CheckforAmp = toCheck
    End If
    
End Function

Public Function GetCreationDateTime() As String
    GetCreationDateTime = Format$(Now, "yyyy-mm-dd") & "T" & Format$(Now, "hh:nn:ss")
End Function


Public Function CreateGUID() As String
    Dim g As GUID
    Dim buffer As String
    buffer = String$(39, vbNullChar)

    If CoCreateGuid(g) = 0 Then
        Call StringFromGUID2(g, StrPtr(buffer), 39)
        CreateGUID = Left$(buffer, InStr(buffer, vbNullChar) - 1)
    Else
        CreateGUID = ""
    End If
End Function


Private Function SafeGetNodeText(parentNode As Object, xpath As String, ns As String) As String
    On Error Resume Next ' Handle errors gracefully
    Dim node As Object
    Set node = parentNode.selectSingleNode(xpath & "[" & ns & "]")
    If Not node Is Nothing Then
        SafeGetNodeText = node.text
    Else
        SafeGetNodeText = ""
    End If
    On Error GoTo 0 ' Reset error handling
End Function

Private Function SafeGetAttribute(parentNode As Object, xpath As String, attributeName As String, ns As String) As String
    On Error Resume Next ' Handle errors gracefully
    Dim node As Object
    Set node = parentNode.selectSingleNode(xpath & "[" & ns & "]")
    If Not node Is Nothing Then
        SafeGetAttribute = node.Attributes.getNamedItem(attributeName).text
    Else
        SafeGetAttribute = ""
    End If
    On Error GoTo 0 ' Reset error handling
End Function

Public Function NoPdfPeppolViewer(filePath As String) As Boolean

    Dim xml As Object  ' MSXML2.DOMDocument
    Dim xsl As Object  ' MSXML2.DOMDocument
    Dim result As String
    Dim utf8Text As String
    utf8Text = MarReadUtf8File(filePath)
    
    NoPdfPeppolViewer = False
    
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False
    xml.validateOnParse = False
    
    'xml.Load filePath

    If Not xml.loadXML(utf8Text) Then
        MsgBox "XML parse error: " & xml.parseError.reason
        Exit Function
    End If

    'If xml.parseError.errorCode <> 0 Then
    '    MsgBox "XML error: " & xml.parseError.reason
    '    Exit Function
    'End If

    Set xsl = CreateObject("MSXML2.DOMDocument.6.0")
    xsl.async = False
    xsl.Load PROGRAM_LOCATION & "xml-templates\peppol\peppol-invoice.xslt"

    If xsl.parseError.errorCode <> 0 Then
        MsgBox "XSLT error: " & xsl.parseError.reason
        Exit Function
    End If

    result = xml.transformNode(xsl)
    Call MarWriteUtf8File(LOCATION_COMPANYDATA & "peppol\in\invoiceNoPdf.html", result)

    ' Save result to HTML
    'Dim f As Integer
    'f = FreeFile
    'Open LOCATION_COMPANYDATA & "peppol\in\invoiceNoPdf.html" For Output As #f
    'Print #f, result
    'Close #f
    NoPdfPeppolViewer = True

End Function

Public Function PeppolHasPdfAttachment(xmlPath As String) As Boolean
    Dim xDoc As New MSXML2.DOMDocument60
    Dim nodeList As MSXML2.IXMLDOMNodeList
    Dim node As MSXML2.IXMLDOMNode

    xDoc.async = False
    xDoc.validateOnParse = False

    If Not xDoc.Load(xmlPath) Then
        Err.Raise vbObjectError + 1, , "Invalid XML: " & xDoc.parseError.reason
    End If

    ' Namespace-aware search
    xDoc.SetProperty "SelectionNamespaces", _
        "xmlns:cbc='urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2' " & _
        "xmlns:cac='urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'"

    Set nodeList = xDoc.selectNodes("//cac:AdditionalDocumentReference/cac:Attachment/cbc:EmbeddedDocumentBinaryObject")

    For Each node In nodeList
        Dim mime As String
        mime = LCase$(node.Attributes.getNamedItem("mimeCode").text)

        If mime = "application/pdf" Then
            PeppolHasPdfAttachment = True
            Exit Function
        End If
    Next node

    PeppolHasPdfAttachment = False
End Function

' Add reference to Microsoft XML, v6.0 (Project > References)
Public Sub ReadUblDocument(filePath As String, showMessageBox As Boolean)
    
    Dim Msg As String
    Dim valueToUse As String
    Dim documentId As String
                
    Dim xmlDoc As Object
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    Dim percentNode As MSXML2.IXMLDOMNode
    
    xmlDoc.async = False
    xmlDoc.Load filePath
    
    uitwisselingOMS = ""
    uitwisselingDATA = ""

    ' Namespace manager (set for XPath queries)
    xmlDoc.SetProperty "SelectionNamespaces", _
        "xmlns:cbc='urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2' " & _
        "xmlns:cac='urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2'"

    Dim sb As String
    sb = ""
    
    'UBL Version
    '1
    uitwisselingOMS = "UBL VersionID"
    uitwisselingDATA = GetNodeText(xmlDoc, "//cbc:UBLVersionID")
    sb = sb & "UBL VersionID: " & GetNodeText(xmlDoc, "//cbc:UBLVersionID") & vbCrLf

    ' Document ID
    valueToUse = GetNodeText(xmlDoc, "//cbc:ID")
    documentId = valueToUse
    '2
    uitwisselingOMS = uitwisselingOMS & vbTab & "documentIdToCheck"
    uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
    sb = sb & "Document ID: " & valueToUse & vbCrLf

    ' IssueDate
    valueToUse = GetNodeText(xmlDoc, "//cbc:IssueDate")
    '3
    uitwisselingOMS = uitwisselingOMS & vbTab & "dateSellerDocumentToCheck"
    uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
    sb = sb & "IssueDate: " & valueToUse & vbCrLf

    ' DueDate
    valueToUse = GetNodeText(xmlDoc, "//cbc:DueDate")
    '4
    uitwisselingOMS = uitwisselingOMS & vbTab & "dateExpiringDocumentToCheck"
    uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
    sb = sb & "DueDate: " & valueToUse & vbCrLf

Dim InvoiceOrCreditnoteTypeCode As String
InvoiceOrCreditnoteTypeCode = "//cbc:InvoiceTypeCode"

Dim isInvoiceTocheck As Boolean
    
TRYFORCREDITNOTE:
    ' DocumentTypeCode
    Dim invTypeNode As Object
    Set invTypeNode = xmlDoc.selectSingleNode(InvoiceOrCreditnoteTypeCode)
    
    If Not invTypeNode Is Nothing Then
        Select Case Dec(Val(invTypeNode.text), "000")
            Case "071", "084"
                isInvoiceTocheck = True
            Case "380", "386"
                isInvoiceTocheck = True
            Case "381"
                isInvoiceTocheck = False
            Case Else
                MsgBox "Onbekende verwerkingscode " & Dec(Val(invTypeNode.text), "000") & vbCrLf & vbCrLf & "Bezorg ons het document. Dank voor medewerking", vbInformation
        End Select
        
        '5
        uitwisselingOMS = uitwisselingOMS & vbTab & "documentTypeCode"
        uitwisselingDATA = uitwisselingDATA & vbTab & Dec(Val(invTypeNode.text), "000")
        sb = sb & "documentTypeCode: " & invTypeNode.text & vbCrLf
        
        If Not invTypeNode.Attributes Is Nothing Then
            Dim listID As String
            On Error Resume Next
            listID = invTypeNode.Attributes.getNamedItem("listID").text
            If listID = "" Then listID = "not found"
            On Error GoTo 0
            sb = sb & "document listID: " & listID & vbCrLf
        End If
    Else
        'MsgBox "InvoiceTypeCode element not found."
        InvoiceOrCreditnoteTypeCode = "//cbc:CreditNoteTypeCode"
        GoTo TRYFORCREDITNOTE
    End If

    ' OrderReference
    Dim orderList As Object
    Set orderList = xmlDoc.selectNodes("//cac:OrderReference")
    If Not orderList Is Nothing Then
        Dim i As Integer
        For i = 0 To orderList.Length - 1
            Dim orderId As String
            orderId = GetNodeText(orderList.Item(i), "cbc:ID")
            If orderId = "" Then orderId = "Order ID: not available"
            sb = sb & "Order ID: " & orderId & vbCrLf
        Next
    End If
    If showMessageBox Then
        MsgBox sb, vbInformation, "Testing UBL DATA versie 0.01"
    End If

    ' Supplier info
    Dim supplierNode As Object
    Set supplierNode = xmlDoc.selectSingleNode("//cac:AccountingSupplierParty/cac:Party")
    If Not supplierNode Is Nothing Then
        Msg = "Supplier info" & vbCrLf & "-------------" & vbCrLf
        
        valueToUse = GetNodeText(supplierNode, "cbc:EndpointID")
        If (Len(valueToUse) = 12 And InStr(valueToUse, "BE")) Then
            valueToUse = Mid(valueToUse, 3)
        End If
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierCompanyIdToCheck"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "endpointOndernemingsnummer " & valueToUse & vbCrLf
        
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierID"
        uitwisselingDATA = uitwisselingDATA & vbTab & GetNodeText(supplierNode, "cac:PartyIdentification/cbc:ID")
        Msg = Msg & "supplierID: " & GetNodeText(supplierNode, "cac:PartyIdentification/cbc:ID") & vbCrLf
        
        'valueToUse = GetNodeText(supplierNode, "cac:PartyName/cbc:Name")
        'If valueToUse = "" Then
            valueToUse = GetNodeText(supplierNode, "cac:PartyLegalEntity/cbc:RegistrationName")
        'End If
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierNameToCheck"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "tradingName: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(supplierNode, "cac:PostalAddress/cbc:StreetName")
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierStreetToCheck"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "street: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(supplierNode, "cac:PostalAddress/cbc:CityName")
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierCityToCheck"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "city: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(supplierNode, "cac:PostalAddress/cbc:PostalZone")
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierPostalCodeToCheck"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "postalZone: " & valueToUse & vbCrLf
                
        valueToUse = GetNodeText(supplierNode, "cac:PostalAddress/cac:Country/cbc:IdentificationCode")
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierCountryCodeToCheck"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "countryCode: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(supplierNode, "cac:PartyTaxScheme/cbc:CompanyID")
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierVatNumberToCheck"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "vatNumber: " & valueToUse & vbCrLf
        If showMessageBox Then
            MsgBox Msg, vbInformation, "Testing UBL DATA versie 0.01"
        End If
    Else
        MsgBox "No AccountingSupplierParty element found."
    End If

    ' PaymentMeans
    Dim pmNodes As Object
    Set pmNodes = xmlDoc.selectNodes("//cac:PaymentMeans")
    If Not pmNodes Is Nothing And pmNodes.Length > 0 Then
        Msg = "PaymentMeans" & vbCrLf & "------------" & vbCrLf
        For i = 0 To pmNodes.Length - 1
            Dim pmNode As Object
            Set pmNode = pmNodes.Item(i)
            
            '1
            uitwisselingOMS = uitwisselingOMS & vbTab & "paymentMeansCode"
            uitwisselingDATA = uitwisselingDATA & vbTab & GetNodeText(pmNode, "cbc:PaymentMeansCode")
            Msg = Msg & "PaymentMeansCode: " & GetNodeText(pmNode, "cbc:PaymentMeansCode") & vbCrLf
            
            valueToUse = GetNodeText(pmNode, "cbc:PaymentID")
            If valueToUse = "" Then
                valueToUse = documentId
            End If
            If InStr(valueToUse, "+") Or InStr(valueToUse, "/") Then
                valueToUse = Replace(valueToUse, "+", "")
                valueToUse = Replace(valueToUse, "/", "")
            End If
            '2
            uitwisselingOMS = uitwisselingOMS & vbTab & "payReferenceToCheck"
            uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
            Msg = Msg & "PaymentID: " & valueToUse & vbCrLf
            
            valueToUse = GetNodeText(pmNode, "cac:PayeeFinancialAccount/cbc:ID")
            '3
            uitwisselingOMS = uitwisselingOMS & vbTab & "supplierIBANToCheck"
            uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
            Msg = Msg & "Payee IBAN: " & valueToUse & vbCrLf
            
            valueToUse = GetNodeText(pmNode, "cac:PayeeFinancialAccount/cbc:Name")
            '4
            uitwisselingOMS = uitwisselingOMS & vbTab & "paySupplierNameToCheck"
            uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
            Msg = Msg & "Account Name: " & valueToUse & vbCrLf
            
            valueToUse = GetNodeText(pmNode, "cac:PayeeFinancialAccount/cac:FinancialInstitutionBranch/cbc:ID")
            '5
            uitwisselingOMS = uitwisselingOMS & vbTab & "supplierBICToCheck"
            uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
            Msg = Msg & "BIC/Branch ID: " & valueToUse & vbCrLf
            
            Msg = Msg & vbCrLf
            Msg = Msg & "Card account (if present)" & vbCrLf
            Msg = Msg & "Card Account ID: " & GetNodeText(pmNode, "cac:CardAccount/cbc:ID") & vbCrLf
            Msg = Msg & "Card Account Name: " & GetNodeText(pmNode, "cac:CardAccount/cbc:Name") & vbCrLf
            Msg = Msg & vbCrLf
            Msg = Msg & "Direct debit mandate (if present)" & vbCrLf
            
            valueToUse = GetNodeText(pmNode, "cac:PaymentMandate/cbc:ID")
            '6
            uitwisselingOMS = uitwisselingOMS & vbTab & "supplierDomMandate"
            uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
            Msg = Msg & "Mandate ID: " & valueToUse & vbCrLf
            
            valueToUse = GetNodeText(pmNode, "cac:PaymentMandate/cbc:PaymentMandateDate")
            '7
            uitwisselingOMS = uitwisselingOMS & vbTab & "supplierDomDate"
            uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
            Msg = Msg & "Mandate Date: " & valueToUse & vbCrLf
        Next
        If showMessageBox Then
            MsgBox Msg, vbInformation, "Testing UBL DATA versie 0.01"
        End If
    Else
        'MsgBox "No PaymentMeans element found."
        '1
        uitwisselingOMS = uitwisselingOMS & vbTab & "paymentMeansCode"
        uitwisselingDATA = uitwisselingDATA & vbTab & "" 'uitwisselingDATA & vbTab & GetNodeText(pmNode, "cbc:PaymentMeansCode")
        '2
        uitwisselingOMS = uitwisselingOMS & vbTab & "payReferenceToCheck"
        uitwisselingDATA = uitwisselingDATA & vbTab & "" 'uitwisselingDATA & vbTab & valueToUse
        '3
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierIBANToCheck"
        uitwisselingDATA = uitwisselingDATA & vbTab & "" 'uitwisselingDATA & vbTab & valueToUse
        '4
        uitwisselingOMS = uitwisselingOMS & vbTab & "paySupplierNameToCheck"
        uitwisselingDATA = uitwisselingDATA & vbTab & "" 'uitwisselingDATA & vbTab & valueToUse
        '5
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierBICToCheck"
        uitwisselingDATA = uitwisselingDATA & vbTab & ""
        '6
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierDomMandate"
        uitwisselingDATA = uitwisselingDATA & vbTab & ""
        '7
        uitwisselingOMS = uitwisselingOMS & vbTab & "supplierDomDate"
        uitwisselingDATA = uitwisselingDATA & vbTab & "" 'uitwisselingDATA & vbTab & valueToUse
    End If

    ' TaxTotal
    Dim msgTax As String
    msgTax = "TaxTotal" & vbCrLf & "--------" & vbCrLf
    Dim taxAmountEl As Object
    Set taxAmountEl = xmlDoc.selectSingleNode("//cac:TaxTotal/cbc:TaxAmount")
    Dim currencyID As String
    currencyID = ""
    If Not taxAmountEl Is Nothing Then
        On Error Resume Next
        currencyID = taxAmountEl.Attributes.getNamedItem("currencyID").text
        On Error GoTo 0
        If currencyID = "" Then
            MsgBox "Attribute currencyID is missing on <cbc:TaxAmount>"
        End If
    End If

    Dim taxTotals As Object
    Set taxTotals = xmlDoc.selectNodes("//cac:TaxTotal")
    If Not taxTotals Is Nothing Then
        For i = 0 To taxTotals.Length - 1
            Dim taxTotalElem As Object
            Set taxTotalElem = taxTotals.Item(i)
            Dim ttAmount As String
            ttAmount = GetNodeText(taxTotalElem, "cbc:TaxAmount")
            msgTax = msgTax & "TaxTotal: " & ttAmount & " " & currencyID & vbCrLf

            Dim subtotals As Object
            Set subtotals = taxTotalElem.selectNodes("cac:TaxSubtotal")
            Dim j As Integer
            For j = 0 To subtotals.Length - 1
                Dim subElem As Object
                Set subElem = subtotals.Item(j)
                msgTax = msgTax & vbCrLf & "SubDetail" & vbCrLf
                msgTax = msgTax & "TaxableAmount: " & GetNodeText(subElem, "cbc:TaxableAmount") & vbCrLf
                msgTax = msgTax & "TaxAmount: " & GetNodeText(subElem, "cbc:TaxAmount") & vbCrLf
                msgTax = msgTax & "Percent: " & GetNodeText(subElem, "cac:TaxCategory/cbc:Percent") & "%" & vbCrLf & vbCrLf
            Next
        Next
        If showMessageBox Then
            MsgBox msgTax, vbInformation, "Testing UBL DATA versie 0.01"
        End If
    End If

    ' LegalMonetaryTotal
    Dim msgMoney As String
    msgMoney = "LegalMonetaryTotal" & vbCrLf & "------------------" & vbCrLf
    Dim moneyTotalEl As Object
    Set moneyTotalEl = xmlDoc.selectSingleNode("//cac:LegalMonetaryTotal")
    If Not moneyTotalEl Is Nothing Then
        msgMoney = msgMoney & "LineExtensionAmount: " & GetNodeText(moneyTotalEl, "cbc:LineExtensionAmount") & vbCrLf
                    
        valueToUse = GetNodeText(moneyTotalEl, "cbc:TaxExclusiveAmount")
        uitwisselingOMS = uitwisselingOMS & vbTab & "totalExclusiveVAT"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        msgMoney = msgMoney & "TaxExclusiveAmount: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(moneyTotalEl, "cbc:TaxInclusiveAmount")
        uitwisselingOMS = uitwisselingOMS & vbTab & "totalInclusiveVAT"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        msgMoney = msgMoney & "TaxInclusiveAmount: " & valueToUse & vbCrLf
        
        msgMoney = msgMoney & "PayableAmount: " & GetNodeText(moneyTotalEl, "cbc:PayableAmount") & " (" & currencyID & ")" & vbCrLf
        If showMessageBox Then
            MsgBox msgMoney, vbInformation, "Testing UBL DATA versie 0.01"
        End If
    End If
    
    ' Customer info
    Dim custNode As Object
    Set custNode = xmlDoc.selectSingleNode("//cac:AccountingCustomerParty")
    If Not custNode Is Nothing Then
        Msg = "Customer info" & vbCrLf & "-------------" & vbCrLf
        
        valueToUse = GetNodeText(custNode, "cbc:CustomerAssignedAccountID")
        uitwisselingOMS = uitwisselingOMS & vbTab & "customerAccountID"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "custAssignedAccountID: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(custNode, "cac:Party/cbc:EndpointID")
        If (Len(valueToUse) = 12 And InStr(valueToUse, "BE")) Then
            valueToUse = Mid(valueToUse, 3)
        End If
        uitwisselingOMS = uitwisselingOMS & vbTab & "customerEndpointID"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "custEndpointID: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(custNode, "cac:Party/cac:PartyName/cbc:Name")
        uitwisselingOMS = uitwisselingOMS & vbTab & "customerName"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "custName: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(custNode, "cac:Party/cac:PostalAddress/cbc:StreetName")
        uitwisselingOMS = uitwisselingOMS & vbTab & "customerStreetName"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "custStreet: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(custNode, "cac:Party/cac:PostalAddress/cbc:CityName")
        uitwisselingOMS = uitwisselingOMS & vbTab & "customerCityName"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "custCity: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(custNode, "cac:Party/cac:PostalAddress/cbc:PostalZone")
        uitwisselingOMS = uitwisselingOMS & vbTab & "customerPostalZone"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "custPostalZone: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(custNode, "cac:Party/cac:PostalAddress/cac:Country/cbc:IdentificationCode")
        uitwisselingOMS = uitwisselingOMS & vbTab & "customerCountryCode"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "custCountryCode: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(custNode, "cac:Party/cac:PartyTaxScheme/cbc:CompanyID")
        uitwisselingOMS = uitwisselingOMS & vbTab & "customerTaxID"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "custTaxID: " & valueToUse & vbCrLf
        
        valueToUse = GetNodeText(custNode, "cac:Party/cac:PartyTaxScheme/cac:TaxScheme/cbc:ID")
        uitwisselingOMS = uitwisselingOMS & vbTab & "customerTaxScheme"
        uitwisselingDATA = uitwisselingDATA & vbTab & valueToUse
        Msg = Msg & "custTaxScheme: " & valueToUse & vbCrLf
        If showMessageBox Then
            MsgBox Msg, vbInformation, "Testing UBL DATA versie 0.01"
        End If
    Else
        MsgBox "No AccountingCustomerParty element found."
    End If

    ' DocumentLines
    documentLinesOMS = ""
    documentLinesDATA = ""

    Dim msgLines As String
    msgLines = ""
    Dim invoiceLines As Object
    If isInvoiceTocheck = True Then
        Set invoiceLines = xmlDoc.selectNodes("//cac:InvoiceLine")
    Else
        Set invoiceLines = xmlDoc.selectNodes("//cac:CreditNoteLine")
    End If
    Dim itemNode As Object
    
    documentLinesOMS = "LineID" & vbTab & "StandardItemID" & vbTab & "SellerID" & vbTab & "Description" & vbTab & "Name" & vbTab & "Quantity" & vbTab & "PriceAmount" & vbTab & "ExtensionAmount" & vbTab & "TaxPercentage"
    If Not invoiceLines Is Nothing Then
        For i = 0 To invoiceLines.Length - 1
            Dim lineNode As Object
            Set lineNode = invoiceLines.Item(i)
            Dim lineID As String, standardItemID As String, desc As String, nameLine As String, sellerID As String, qty As String, extensionAmount As String, price As String, taxPercentage As String
            
            lineID = GetNodeText(lineNode, ".//cbc:ID")
            If lineID = "" Then lineID = "-"
            
            desc = GetNodeText(lineNode, ".//cbc:Description")
            If desc = "" Then desc = "-"
            
            nameLine = GetNodeText(lineNode, ".//cbc:Name")
            If nameLine = "" Then nameLine = "-"
            
            sellerID = GetNodeText(lineNode, ".//cac:SellersItemIdentification")
            If sellerID = "" Then sellerID = "-"
            
            standardItemID = GetNodeText(lineNode, ".//cac:StandardItemIdentification")
            If standardItemID = "" Then standardItemID = "-"
            
            If isInvoiceTocheck = True Then
                qty = GetNodeText(lineNode, ".//cbc:InvoicedQuantity")
            Else
                qty = GetNodeText(lineNode, ".//cbc:CreditedQuantity")
            End If
            
            
            If qty = "" Then qty = "-"
            
            extensionAmount = GetNodeText(lineNode, ".//cbc:LineExtensionAmount")
            If Val(extensionAmount) = 0 Then
                extensionAmount = "-"
            End If
            
            price = GetNodeText(lineNode, ".//cbc:PriceAmount")
            If Val(price) = 0 Then
                price = "-"
            End If
            
            taxPercentage = "0"
            ' Get the Percent node under ClassifiedTaxCategory
            Set percentNode = lineNode.selectSingleNode("cac:Item/cac:ClassifiedTaxCategory/cbc:Percent")
        
            If Not percentNode Is Nothing Then
                taxPercentage = percentNode.text
            Else
                taxPercentage = "0"
            End If
            
            If extensionAmount = "-" Then
            Else
                documentLinesDATA = documentLinesDATA & lineID & vbTab & standardItemID & vbTab & sellerID & vbTab & desc & vbTab & nameLine & vbTab & qty & vbTab & price & vbTab & extensionAmount & vbTab & taxPercentage & vbCrLf
            End If
            msgLines = msgLines & "Item: " & desc & ", Quantity: " & qty & ", Price: " & price & vbCrLf
        Next
        If msgLines <> "" Then
            If showMessageBox Then
                MsgBox msgLines, vbInformation, "Testing UBL DATA versie 0.01"
            End If
        End If
    End If
    documentLinesDATA = documentLinesOMS & vbCrLf & documentLinesDATA
    '? invoicelinesdata
    
End Sub

Private Function GetSupplierNode(xmlDoc As MSXML2.DOMDocument60) As MSXML2.IXMLDOMNode
    Dim node As MSXML2.IXMLDOMNode
    Set node = xmlDoc.selectSingleNode("//cac:AccountingSupplierParty/cac:Party")
    
    If node Is Nothing Then
        Err.Raise vbObjectError + 2, , "SupplierParty node not found"
    End If
    
    Set GetSupplierNode = node
End Function

Private Function GetPaymentMeansNodes(xmlDoc As MSXML2.DOMDocument60) As MSXML2.IXMLDOMNodeList
    Set GetPaymentMeansNodes = xmlDoc.selectNodes("//cac:PaymentMeans")
    
    If GetPaymentMeansNodes.Length = 0 Then
        Err.Raise vbObjectError + 2, , "No PaymentMeans elements found"
    End If
End Function

Private Function NodeText(parentNode As MSXML2.IXMLDOMNode, xpath As String) As String
    Dim child As MSXML2.IXMLDOMNode
    Set child = parentNode.selectSingleNode(xpath)
    
    If Not child Is Nothing Then
        NodeText = Trim(child.text)
    Else
        NodeText = ""
    End If
End Function

' Helper for safe node text extraction
Private Function GetNodeText(parentNode As Object, xpath As String) As String
    Dim node As Object
    Set node = parentNode.selectSingleNode(xpath)
    If Not node Is Nothing Then
        GetNodeText = Trim(node.text)
    Else
        GetNodeText = ""
    End If
End Function

Public Function CheckPeppolRegistration(peppolID As String) As String
    
    'r = ShellExecute(0, "open", "https://directory.peppol.eu/search/1.0/json?q=iso6523-actorid-upis:" + Me.tbPeppolID, 0, 0, 1)
    
    Dim http As Object
    Dim xmlDoc As Object
    Dim docNodes As Object
    Dim docNode As Object
    Dim entityNode As Object
    
    Dim url As String
    Dim responseText As String
    responseText = ""

    url = "https://directory.peppol.eu/search/1.0/json?q=iso6523-actorid-upis:" + peppolID
    On Local Error Resume Next
    Err = 0

    Screen.MousePointer = vbHourglass
    DoEvents
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send

    If http.Status = 200 Then
        responseText = http.responseText
    Else
        'MsgBox "HTTP Error: " & http.Status
        responseText = ""
    End If
    Screen.MousePointer = vbNormal
    DoEvents
    
    Set http = Nothing
    Set xmlDoc = Nothing
    Set docNodes = Nothing
    Set docNode = Nothing
    Set entityNode = Nothing
    CheckPeppolRegistration = responseText

End Function


Public Sub ExtractPdfAttachments(ByVal ublFilePath As String, ByVal xmlLocation As String)

    Dim xml As New MSXML2.DOMDocument60
    xml.async = False
    xml.preserveWhiteSpace = True

    If Not xml.Load(ublFilePath) Then
        MsgBox "XML kon niet geladen worden: " & xml.parseError.reason
        Exit Sub
    End If

    ' Correct namespace registration for VB6
    xml.SetProperty "SelectionNamespaces", _
        "xmlns:cac='urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2' " & _
        "xmlns:cbc='urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2'"

    ' Find ALL attachments
    Dim nodes As MSXML2.IXMLDOMNodeList
    Set nodes = xml.selectNodes("//cac:AdditionalDocumentReference/cac:Attachment/cbc:EmbeddedDocumentBinaryObject")

    If nodes Is Nothing Or nodes.Length = 0 Then
        MsgBox "Geen bijlagen gevonden."
        Exit Sub
    End If

    Dim node As MSXML2.IXMLDOMNode
    For Each node In nodes

        Dim filename As String
        Dim mime As String

        filename = node.Attributes.getNamedItem("filename").text
        mime = node.Attributes.getNamedItem("mimeCode").text

        If LCase$(mime) <> "application/pdf" Then GoTo NextNode

        Dim base64 As String
        base64 = CleanBase64(node.text)

        Dim bytes() As Byte
        bytes = Base64Decode(base64)

        Dim pdfPath As String
        pdfPath = Left$(ublFilePath, Len(ublFilePath) - 4) & "_" & filename

        SaveBinary pdfPath, bytes
        DoEvents
        
        Dim lookForPDF As String
        'look for documents to print or send with peppol and remove them
        On Error Resume Next
        Err = 0
        lookForPDF = Dir(pdfPath)
        DoEvents
        If lookForPDF = "" Then
            MsgBox "Er is geen PDF beschikbaar in " & vbCrLf & lookForPDF & vbCrLf & vbCrLf & "Opteer XML tonen", vbExclamation
        Else
            If ShellExecuteWithFallback(xmlLocation + lookForPDF) = False Then
                MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
            End If
        End If
        

        'ShellExecute 0, "open", pdfPath, vbNullString, vbNullString, 1

NextNode:
    Next node

End Sub


Public Function CleanBase64(ByVal s As String) As String
    Dim tmp As String
    tmp = Replace(s, vbCr, "")
    tmp = Replace(tmp, vbLf, "")
    tmp = Replace(tmp, vbTab, "")
    tmp = Replace(tmp, " ", "")
    CleanBase64 = tmp
End Function

Public Function Base64Decode(ByVal base64String As String) As Byte()
    Dim xml As Object
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")

    Dim node As Object
    Set node = xml.createElement("b64")

    node.dataType = "bin.base64"
    node.text = base64String

    Base64Decode = node.nodeTypedValue
End Function

Public Sub SaveBinary(ByVal filePath As String, ByRef bytes() As Byte)
    Dim hFile As Integer
    hFile = FreeFile
    Open filePath For Binary Access Write As #hFile
        Put #hFile, , bytes
    Close #hFile
End Sub


' Reference: Microsoft ActiveX Data Objects 2.x Library
Function MarReadUtf8File(filename As String) As String
    
    Dim available As String
    available = Dir(filename)
    If available = "" Then
        MarReadUtf8File = ""
        Exit Function
    End If
    
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text
    stream.Charset = "utf-8"
    On Error Resume Next
    stream.Open
    stream.LoadFromFile filename
    
    MarReadUtf8File = stream.ReadText
    stream.Close
    Set stream = Nothing
End Function

' Reference: Microsoft ActiveX Data Objects 2.x Library (Project > References)
Sub MarWriteUtf8File(filename As String, text As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile filename, 2 ' 2 = adSaveCreateOverWrite
    stream.Close
    Set stream = Nothing
End Sub
