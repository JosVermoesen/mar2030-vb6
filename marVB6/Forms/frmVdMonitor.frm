VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FormB2BVMonitor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Monitor B2B Verkoopdocumenten"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9675
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ButtonResponsesToClient 
      Caption         =   "Reacties (van klant)"
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
      Left            =   5880
      TabIndex        =   9
      Top             =   6720
      Width           =   2295
   End
   Begin VB.CommandButton ButtonSentReceiptClient 
      Caption         =   "Ontvangstbewijs (server klant)"
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
      Left            =   2880
      TabIndex        =   7
      Top             =   6720
      Width           =   2895
   End
   Begin VB.CommandButton ButtonSentReceiptSeller 
      Caption         =   "Verzend ID"
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
      Left            =   1440
      TabIndex        =   6
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton ButtonShowToSendPDF 
      Caption         =   "PDF Tonen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton ButtonShowPDFSent 
      Caption         =   "PDF Tonen"
      Enabled         =   0   'False
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton ButtonClose 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   8280
      TabIndex        =   0
      Top             =   6720
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgToSend 
      Height          =   2775
      Left            =   120
      TabIndex        =   2
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgSent 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
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
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Reeds verzonden"
      Height          =   255
      Left            =   8040
      TabIndex        =   8
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label LabelToBook 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nog te verzenden (MarSync)"
      Height          =   255
      Left            =   7320
      TabIndex        =   3
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "FormB2BVMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim fileUrlAsPdf As String
Dim selectedFile As String
    
Dim rsSellerUBL As New ADODB.Recordset
Function GetSentReceipt(vDocument As String, asSeller As Integer) As String

    GetSentReceipt = ""
            
    Set rsSellerUBL = New ADODB.Recordset
    On Error Resume Next
    Err = 0
    rsSellerUBL.CursorLocation = adUseClient
                
    Dim sSQL As String
        
    sSQL = "SELECT "
    sSQL = sSQL + "Dokumenten.v033, Dokumenten.v405, Dokumenten.v406, Dokumenten.v408 "
    sSQL = sSQL + "FROM  Dokumenten "
    sSQL = sSQL + "WHERE Dokumenten.v033 = '" + vDocument + "' "
       
    Screen.MousePointer = vbHourglass
    rsSellerUBL.Open sSQL, adntDB, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        GetSentReceipt = Err.Description
    ElseIf rsSellerUBL.RecordCount = 0 Then
        GetSentReceipt = "Niets gevonden"
    ElseIf rsSellerUBL.RecordCount = 1 Then
        'check status
        Dim sentID As String
        If asSeller = 1 Then
            sentID = Trim(rsSellerUBL("v405"))
            If sentID = "" Then
                GetSentReceipt = "Leeg"
            Else
                GetSentReceipt = sentID
            End If
        ElseIf asSeller = 0 Then
            sentID = Trim(rsSellerUBL("v406"))
            If sentID = "" Then
                GetSentReceipt = "Leeg"
            Else
                GetSentReceipt = sentID
            End If
        ElseIf asSeller = 2 Then
            sentID = Trim(rsSellerUBL("v408"))
            If sentID = "" Then
                GetSentReceipt = "Leeg"
            Else
                GetSentReceipt = sentID
            End If
        End If
    End If

End Function

Private Sub ButtonClose_Click()
    
    Unload Me
    
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
            stream.SaveToFile Mid(ublFileUrl, 1, Len(ublFileUrl) - 4) & ".pdf", 2 'adSaveCreateOverWrite
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



Private Sub ButtonResponsesToClient_Click()

    Dim selectedRowItem As String
    Dim result As String
        
    mfgSent.Col = 5
    selectedRowItem = mfgSent.text
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

Private Sub ButtonSentReceiptClient_Click()

    Dim selectedRowItem As String
    Dim result As String
    
    mfgSent.Col = 5
    selectedRowItem = mfgSent.text
    If selectedRowItem = "" Then
    Else
        result = GetSentReceipt(Mid(selectedRowItem, 1, 11), 0)
        If InStr(result, """count""" + ": 0") Then
            MsgBox "Nog te bevestigen. Vernieuw met MarSync", vbInformation
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
    
    mfgSent.Col = 5
    selectedRowItem = mfgSent.text
    If selectedRowItem = "" Then
    Else
        result = GetSentReceipt(Mid(selectedRowItem, 1, 11), 1)
        Load FormReactionsDialog
        FormReactionsDialog.TextBoxReactions.text = result
        FormReactionsDialog.Caption = "Verzend ID"
        FormReactionsDialog.Show 1
    End If

End Sub

Private Sub ButtonShowPDFSent_Click()

    Dim selectedRowItem As String
    
    mfgSent.Col = 5
    selectedRowItem = mfgSent.text
    If selectedRowItem = "" Then
    Else
        selectedFile = LOCATION_COMPANYDATA & "peppol\out\" & selectedRowItem
        ExtractPdfFromUBLDocument (selectedFile)
        fileUrlAsPdf = Left(selectedFile, Len(selectedFile) - 3) + "pdf"
        ' Voorbeeld gebruik
        If ShellExecuteWithFallback(fileUrlAsPdf) = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
    End If

End Sub


Private Sub ButtonShowToSendPDF_Click()

    Dim selectedRowItem As String
    
    mfgToSend.Col = 5
    selectedRowItem = mfgToSend.text
    If selectedRowItem = "" Then
    Else
        selectedFile = LOCATION_COMPANYDATA & "peppol\out\" & selectedRowItem
        ExtractPdfFromUBLDocument (selectedFile)
        fileUrlAsPdf = Left(selectedFile, Len(selectedFile) - 3) + "pdf"
        ' Voorbeeld gebruik
        If ShellExecuteWithFallback(fileUrlAsPdf) = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
    End If

End Sub


Private Sub mfgSent_Click()

    Me.ButtonShowToSendPDF.Enabled = False
    If mfgSent.Rows = 2 Then
        Me.ButtonShowPDFSent.Enabled = False
    Else
        Me.ButtonShowPDFSent.Enabled = True
        Me.ButtonSentReceiptSeller.Enabled = True
        Me.ButtonSentReceiptClient.Enabled = True
        Me.ButtonSentReceiptSeller.Default = True
        Me.ButtonResponsesToClient.Enabled = True
    End If

End Sub

Private Sub mfgSent_GotFocus()

    Dim fileId As String
    fileId = mfgSent.TextMatrix(mfgSent.Row, 5)
    If fileId = "" Then
        Me.ButtonSentReceiptSeller.Default = False
        Me.ButtonSentReceiptSeller.Enabled = False
        Me.ButtonSentReceiptClient.Enabled = False
        Me.ButtonResponsesToClient.Enabled = False
        Me.ButtonShowPDFSent.Enabled = False
    Else
        Me.ButtonSentReceiptSeller.Default = True
        Me.ButtonSentReceiptSeller.Enabled = True
        Me.ButtonSentReceiptClient.Enabled = True
        Me.ButtonResponsesToClient.Enabled = True
        Me.ButtonShowPDFSent.Enabled = True
    End If
    
End Sub

Private Sub mfgSent_RowColChange()

    Dim fileId As String
    fileId = mfgSent.TextMatrix(mfgSent.Row, 5)
    If fileId = "" Then
        Me.ButtonSentReceiptSeller.Default = False
        Me.ButtonSentReceiptSeller.Enabled = False
        Me.ButtonSentReceiptClient.Enabled = False
        Me.ButtonResponsesToClient.Enabled = False
        Me.ButtonShowPDFSent.Enabled = False
    Else
        Me.ButtonSentReceiptSeller.Default = True
        Me.ButtonSentReceiptSeller.Enabled = True
        Me.ButtonSentReceiptClient.Enabled = True
        Me.ButtonResponsesToClient.Enabled = True
        Me.ButtonShowPDFSent.Enabled = True
    End If
    
End Sub

Private Sub mfgToSend_Click()

    Me.ButtonShowPDFSent.Enabled = False
    If mfgToSend.Rows = 2 Then
        Me.ButtonShowToSendPDF.Enabled = False
    Else
        Me.ButtonShowToSendPDF.Enabled = True
    End If

End Sub

Private Sub mfgToSend_GotFocus()

    Me.ButtonSentReceiptSeller.Default = False
    Me.ButtonSentReceiptSeller.Enabled = False
    Me.ButtonSentReceiptClient.Enabled = False
    
End Sub

