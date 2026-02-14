VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FormB2BAMonitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor B2B Aankoopdocumenten"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   9870
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ButtonShowToBookXML 
      Caption         =   "XML Tonen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton ButtonShowBookedXML 
      Caption         =   "XML Tonen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   11
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton ButtonShowPeppolDocTypes 
      Caption         =   "Peppol Document types tonen"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton ButtonSentReceiptSeller 
      Caption         =   "Verzendbewijs"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CommandButton ButtonResponsesToSeller 
      Caption         =   "Reacties (aan leverancier)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   6720
      Width           =   2775
   End
   Begin VB.CommandButton ButtonLoadDocument 
      Caption         =   "Document Inladen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton ButtonShowBookedPDF 
      Caption         =   "PDF Tonen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton ButtonShowToBookPDF 
      Caption         =   "PDF Tonen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton ButtonClose 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgToBook 
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4895
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgBooked 
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4895
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   0
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label LabelBooked 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Ingeboekt"
      Height          =   255
      Left            =   8520
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label LabelToBook 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "In te boeken"
      Height          =   255
      Left            =   8400
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FormB2BAMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim fileUrlAsPdf As String
Dim selectedFile As String

Dim rsBuyerUBL As New ADODB.Recordset
Private Sub ButtonClose_Click()

    uitwisselingOMS = ""
    uitwisselingDATA = ""
    documentLinesOMS = ""
    documentLinesDATA = ""
    Unload Me
    
End Sub


Private Sub ButtonLoadDocument_Click()
    
    Dim selectedRowItem As String
    
    mfgToBook.Col = 5
    selectedRowItem = mfgToBook.text
    If selectedRowItem = "" Then
    Else
        XLogKey = LOCATION_COMPANYDATA & "peppol\in\" & selectedRowItem
        Me.Hide
    End If

End Sub

Private Sub ButtonSentReceiptClient_Click()

End Sub

Private Sub ButtonResponsesToSeller_Click()

    Dim selectedRowItem As String
    Dim result As String
    
    Me.mfgBooked.Col = 1
    selectedRowItem = Me.mfgBooked.text
    If selectedRowItem = "" Then
    Else
        result = GetSentReceipt(Mid(selectedRowItem, 1, 11), 2)
        If InStr(result, """count""" + ": 0") Then
            MsgBox "Geen", vbInformation
        Else
            Load FormReactionsDialog
            FormReactionsDialog.TextBoxReactions.text = result
            FormReactionsDialog.Show 1
        End If
    End If

End Sub

Private Sub ButtonSentReceiptSeller_Click()

    Dim selectedRowItem As String
    Dim result As String
    
    Me.mfgBooked.Col = 1
    selectedRowItem = Me.mfgBooked.text
    If selectedRowItem = "" Then
    Else
        result = GetSentReceipt(Mid(selectedRowItem, 1, 11), 1)
        If InStr(result, """count""" + ": 0") Then
            MsgBox "Nog te bevestigen. Vernieuw met MarSync", vbInformation
        Else
            MsgBox result, vbInformation
        End If
    End If

End Sub

Private Sub ButtonShowBookedXML_Click()

    Dim selectedRowItem As String
        
    selectedRowItem = mfgBooked.TextMatrix(mfgBooked.Row, 5)
    If selectedRowItem = "" Then
    Else
        selectedFile = LOCATION_COMPANYDATA & "peppol\in\" & selectedRowItem
        Dim resultHere As Boolean
        resultHere = NoPdfPeppolViewer(selectedFile)
        If resultHere Then
            'TODO show invoiceNoPdf.html
            If ShellExecuteWithFallback(LOCATION_COMPANYDATA & "peppol\in\invoiceNoPdf.html") = False Then
                MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
            End If
        Else
            MsgBox "Iets ging verkeerd", vbInformation
        End If
    End If

End Sub

Private Sub ButtonShowPeppolDocTypes_Click()

    Dim documentTypesPDF As String
    
    documentTypesPDF = PROGRAM_LOCATION & "Def\PeppolDocTypes.pdf"
    Dim lookForPDF As String
    'look for document
    On Error Resume Next
    Err = 0
    lookForPDF = Dir(documentTypesPDF)
    DoEvents
    If lookForPDF = "" Then
        MsgBox "PDF niet gevonden in " & vbCrLf & documentTypesPDF & vbCrLf & vbCrLf & "Controleer correcte installatie MarIntegraal", vbExclamation
    Else
        If ShellExecuteWithFallback(documentTypesPDF) = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
    End If

End Sub

Private Sub ExtractPdfFromUBLDocument(ublFileUrl As String)
    
    Dim xmlDoc As New MSXML2.DOMDocument
    Dim node As MSXML2.IXMLDOMNode
    Dim base64Data As String
    Dim byteData() As Byte
    Dim stream As Object

    xmlDoc.async = False
    xmlDoc.Load ublFileUrl

    If xmlDoc.parseError.errorCode = 0 Then
        Set node = xmlDoc.selectSingleNode("//cbc:EmbeddedDocumentBinaryObject")
        If Not node Is Nothing Then
            base64Data = node.text
            byteData = Base64Decode_MSXML(base64Data)
            'Base64Decode_MSXML

            Set stream = CreateObject("ADODB.Stream")
            stream.Type = 1 'adTypeBinary
            stream.Open
            stream.Write byteData
            'TO SAVE THE PDF
            On Local Error Resume Next
            stream.SaveToFile Mid(ublFileUrl, 1, Len(ublFileUrl) - 4) & ".pdf", 2 'adSaveCreateOverWrite
            If Err Then
                MsgBox "Wegschrijven PDF is mislukt." & vbCrLf & vbCrLf & "Zelfde bestand staat nog open", vbExclamation
            End If
            
            'TODO: OPEN THE PDF?
            stream.Close
        End If
    End If

End Sub

' Reference: Microsoft XML, v6.0 (or use DOMDocument30/40 if needed)
Public Function Base64Decode_MSXML(ByVal sB64 As String) As Byte()
    Dim dom As MSXML2.DOMDocument60
    Dim el As MSXML2.IXMLDOMElement
    
    Set dom = New MSXML2.DOMDocument60
    Set el = dom.createElement("b64")
    el.dataType = "bin.base64"
    el.text = sB64                ' assign Base64 text
    Base64Decode_MSXML = el.nodeTypedValue ' returns Byte()
End Function

Public Function Base64Encode_MSXML(ByRef bytes() As Byte) As String
    Dim dom As MSXML2.DOMDocument60
    Dim el As MSXML2.IXMLDOMElement
    
    Set dom = New MSXML2.DOMDocument60
    Set el = dom.createElement("b64")
    el.dataType = "bin.base64"
    el.nodeTypedValue = bytes     ' assign bytes
    Base64Encode_MSXML = el.text  ' returns Base64 string
End Function



Private Sub mfgBooked_Click()

    Me.ButtonShowToBookPDF.Enabled = False
    If mfgBooked.Rows = 2 Then
        Me.ButtonShowBookedPDF.Enabled = False
        
        Me.ButtonResponsesToSeller.Enabled = False
        Me.ButtonSentReceiptSeller.Enabled = False
    Else
        Me.ButtonShowBookedPDF.Enabled = True
        Me.ButtonResponsesToSeller.Enabled = True
        Me.ButtonSentReceiptSeller.Enabled = True
    End If

End Sub


Private Sub mfgBooked_GotFocus()

    Me.ButtonLoadDocument.Enabled = False
    
    Me.ButtonShowToBookXML.Enabled = False
    Me.ButtonShowToBookPDF.Enabled = False
    
    Me.ButtonShowBookedXML.Enabled = True
    Me.ButtonShowBookedPDF.Enabled = True
    
End Sub

Private Sub mfgBooked_RowColChange()

    Dim selectedRowItem As String
    selectedRowItem = mfgBooked.TextMatrix(mfgBooked.Row, 5)
    
    If selectedRowItem <> "" Then
        If PeppolHasPdfAttachment(LOCATION_COMPANYDATA & "peppol\in\" & selectedRowItem) Then
            Me.ButtonShowBookedPDF.Visible = True
        Else
            Me.ButtonShowBookedPDF.Visible = False
        End If
    End If
    
    Me.mfgBooked.Col = 1
    selectedRowItem = Me.mfgBooked.text
    If selectedRowItem = "" Then
        Me.ButtonSentReceiptSeller.Visible = False
        Me.ButtonResponsesToSeller.Visible = False
    Else
        Me.ButtonSentReceiptSeller.Visible = True
        Dim result As String
        result = GetSentReceipt(Mid(selectedRowItem, 1, 11), 2)
        If InStr(result, """count""" + ": 0") Then
            ButtonResponsesToSeller.Visible = False
        Else
            ButtonResponsesToSeller.Visible = True
        End If
    End If
        
End Sub

Private Sub mfgToBook_Click()

    Me.ButtonShowBookedPDF.Enabled = False
    If mfgToBook.Rows = 2 Then
        Me.ButtonShowToBookPDF.Enabled = False
        Me.ButtonShowToBookXML.Enabled = False
    Else
        Me.ButtonShowToBookPDF.Enabled = True
        Me.ButtonShowToBookXML.Enabled = True
        Me.ButtonLoadDocument.Enabled = True
        Me.ButtonLoadDocument.Default = True
    End If

End Sub


Private Sub mfgToBook_GotFocus()

    Dim fileId As String
    fileId = mfgToBook.TextMatrix(mfgToBook.Row, 5)
    If fileId = "" Then
        Me.ButtonLoadDocument.Default = False
        Me.ButtonLoadDocument.Enabled = False
    Else
        Me.ButtonLoadDocument.Default = True
        Me.ButtonLoadDocument.Enabled = True
    End If
    If mfgToBook.Rows = 2 Then
    Else
        Me.ButtonShowToBookPDF.Enabled = True
        Me.ButtonShowToBookXML.Enabled = True
    End If
    
    Me.ButtonShowBookedPDF.Enabled = False
    Me.ButtonResponsesToSeller.Enabled = False
    Me.ButtonSentReceiptSeller.Enabled = False

End Sub

Private Sub mfgToBook_LostFocus()

    Me.ButtonLoadDocument.Default = False

End Sub


Private Sub mfgToBook_RowColChange()
    
    Dim fileId As String
    fileId = mfgToBook.TextMatrix(mfgToBook.Row, 5)
    
    If fileId = "" Then
        Me.ButtonLoadDocument.Default = False
        Me.ButtonLoadDocument.Enabled = False
    Else
        Me.ButtonLoadDocument.Default = True
        Me.ButtonLoadDocument.Enabled = True
        
        If PeppolHasPdfAttachment(LOCATION_COMPANYDATA & "peppol\in\" & fileId) Then
            Me.ButtonShowToBookPDF.Visible = True
        Else
            Me.ButtonShowToBookPDF.Visible = False
        End If
    End If
    
End Sub

Function GetSentReceipt(vDocument As String, asSeller As Integer) As String

    GetSentReceipt = ""
            
    Set rsBuyerUBL = New ADODB.Recordset
    On Error Resume Next
    Err = 0
    rsBuyerUBL.CursorLocation = adUseClient
                
    Dim sSQL As String
        
    sSQL = "SELECT "
    sSQL = sSQL + "Dokumenten.v033, Dokumenten.v406, Dokumenten.v408 "
    sSQL = sSQL + "FROM  Dokumenten "
    sSQL = sSQL + "WHERE Dokumenten.v033 = '" + vDocument + "' "
       
    Screen.MousePointer = vbHourglass
    rsBuyerUBL.Open sSQL, adntDB, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        GetSentReceipt = Err.Description
    ElseIf rsBuyerUBL.RecordCount = 0 Then
        GetSentReceipt = "Niets gevonden"
    ElseIf rsBuyerUBL.RecordCount = 1 Then
        'check status
        Dim sentID As String
        If asSeller = 1 Then
            sentID = Trim(rsBuyerUBL("v406"))
            If sentID = "" Then
                GetSentReceipt = "Geen"
            Else
                GetSentReceipt = sentID
            End If
        ElseIf asSeller = 2 Then
            sentID = Trim(rsBuyerUBL("v408"))
            If sentID = "" Then
                GetSentReceipt = "Geen"
            Else
                GetSentReceipt = sentID
            End If
        End If
    End If

End Function

Private Sub ButtonShowToBookPDF_Click()
    
    Dim selectedRowItem As String
    
    mfgToBook.Col = 5
    selectedRowItem = mfgToBook.text
    If selectedRowItem = "" Then
    Else
        selectedFile = LOCATION_COMPANYDATA & "peppol\in\" & selectedRowItem
        ExtractPdfFromUBLDocument (selectedFile)
        fileUrlAsPdf = Left(selectedFile, Len(selectedFile) - 3) + "pdf"
        ' Voorbeeld gebruik
        
        DoEvents
        Dim lookForPDF As String
        'look for documents to print or send with peppol and remove them
        On Error Resume Next
        Err = 0
        lookForPDF = Dir(fileUrlAsPdf)
        DoEvents
        If lookForPDF = "" Then
            MsgBox "Er is geen PDF beschikbaar in " & vbCrLf & selectedFile & vbCrLf & vbCrLf & "Opteer eventueel voor XML tonen in MarSync", vbExclamation
        Else
            If ShellExecuteWithFallback(fileUrlAsPdf) = False Then
                MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
            End If
        End If
    End If
    
End Sub

Private Sub ButtonShowToBookXML_Click()

    'TODO: fix for ButtonShowToBookXML
    'Dim selectedRowItem As String
    'mfgToBook.Col = 5
    'selectedRowItem = mfgToBook.Text
    'If selectedRowItem = "" Then
    'Else
    '    selectedFile = LOCATION_COMPANYDATA & "peppol\in\" & selectedRowItem
    '    Dim resultHere As Boolean
    '    resultHere = NoPdfPeppolViewer(selectedFile)
    'End If
    
    Dim selectedRowItem As String
    
    selectedRowItem = mfgToBook.TextMatrix(mfgToBook.Row, 5)
    If selectedRowItem = "" Then
    Else
        selectedFile = LOCATION_COMPANYDATA & "peppol\in\" & selectedRowItem
        Dim resultHere As Boolean
        resultHere = NoPdfPeppolViewer(selectedFile)
        If resultHere Then
            'TODO show invoiceNoPdf.html
            If ShellExecuteWithFallback(LOCATION_COMPANYDATA & "peppol\in\invoiceNoPdf.html") = False Then
                MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
            End If
        Else
            MsgBox "Iets ging verkeerd", vbInformation
        End If
    End If

End Sub

Private Sub ButtonShowBookedPDF_Click()
    
    Dim selectedRowItem As String
    
    mfgBooked.Col = 5
    selectedRowItem = mfgBooked.text
    If selectedRowItem = "" Then
    Else
        selectedFile = LOCATION_COMPANYDATA & "peppol\in\" & selectedRowItem
        Call ExtractPdfAttachments(selectedFile, LOCATION_COMPANYDATA + "peppol\in\")
        'ExtractPdfFromUBLDocument (selectedFile)
        'fileUrlAsPdf = Left(selectedFile, Len(selectedFile) - 3) + "pdf"
        ' Voorbeeld gebruik
        
        'DoEvents
        'Dim lookForPDF As String
        'look for documents to print or send with peppol and remove them
        'On Error Resume Next
        'Err = 0
        'lookForPDF = Dir(fileUrlAsPdf)
        'DoEvents
        'If lookForPDF = "" Then
        '    MsgBox "Er is geen PDF beschikbaar in " & vbCrLf & selectedFile & vbCrLf & vbCrLf & "Opteer eventueel voor XML tonen in MarSync", vbExclamation
        'Else
        '    If ShellExecuteWithFallback(fileUrlAsPdf) = False Then
        '        MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        '    End If
        'End If
    End If
    
End Sub

