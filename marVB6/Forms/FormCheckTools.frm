VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ValidatingTool 
   Caption         =   "Peppol Tools"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   5940
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TextBoxSupportedDocuments 
      Height          =   495
      Left            =   2400
      TabIndex        =   14
      Top             =   1320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgLijst 
      Height          =   5895
      Left            =   120
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   10398
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   0
      FixedCols       =   0
      FocusRect       =   2
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CheckBox CheckBoxOnlyRecent 
      Caption         =   "met facturatie recente jaren"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   13
      Top             =   120
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton cbCopyToClipBoard 
      Caption         =   "Kopie"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cbCheckAllPartners 
      Caption         =   "Partijen &B2B oplijsten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton cbCheckPeppolRegistration 
      Caption         =   "Controle &Peppol Registratie"
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2400
      Width           =   2175
   End
   Begin VB.TextBox tbPeppolID 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton cbCheckCompanyNumber 
      Caption         =   "&Onderneming Opzoeken"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox tbCompanyNumber 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton cbCheckVatNumber 
      Caption         =   "&Btw Nummer Opzoeken"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox tbVatNumber 
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
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   3255
   End
   Begin VB.CommandButton ButtonClose 
      Cancel          =   -1  'True
      Caption         =   "ButtonClose"
      Height          =   375
      Left            =   1080
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblPeppolId 
      Caption         =   "Peppol ID (Europa)"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label lblCompanyNumber 
      Caption         =   "Ondernemingsnummer (Enkel België!)"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   2775
   End
   Begin VB.Label lblVatNumber 
      Caption         =   "Btw Nummer (Europa)"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "ValidatingTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Option Explicit

Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Dim rsPartners     As New ADODB.Recordset
Dim r As Long
Dim vatList As String
Dim vatListAll As String
Dim partnerIs As String

Private Sub ButtonClose_Click()

    Hide
    
End Sub

Private Sub cbCheckAllPartners_Click()

    If InStr(Me.Caption, "Klanten") Then
        partnerIs = "Klanten"
    Else
        partnerIs = "Leveranciers"
    End If
    
    Set rsPartners = New ADODB.Recordset
            
    On Error Resume Next
    Err = 0
    rsPartners.CursorLocation = adUseClient
    If partnerIs = "Klanten" Then
        Msg = "SELECT DISTINCT Klanten.A110, Klanten.v150, Klanten.A161, Klanten.A100, Klanten.v224 AS [eMail] FROM Klanten, Dokumenten WHERE Dokumenten.v034 = 'K'+Klanten.A110 AND Dokumenten.v033 Like 'V%' AND len(Klanten.A161) > 1"
    Else
        Msg = "SELECT DISTINCT Leveranciers.A110, Leveranciers.v150, Leveranciers.A161, Leveranciers.A100, Leveranciers.v224 AS [eMail] FROM Leveranciers, Dokumenten WHERE Dokumenten.v034 = 'L'+Leveranciers.A110 AND Dokumenten.v033 Like 'A%'"
    End If
    If Me.CheckBoxOnlyRecent.Value = vbChecked Then Msg = Msg & " AND Dokumenten.v035 Like '202%'"
    SnelHelpPrint Msg, BL_LOGGING
    Screen.MousePointer = vbHourglass
    rsPartners.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        Screen.MousePointer = vbNormal
        Exit Sub
    ElseIf rsPartners.RecordCount = 0 Then
        MsgBox "Geen klanten te vinden met btw nummer.", vbInformation
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    
    vatList = ""
    vatListAll = ""
    mfgLijst.Clear
    mfgLijst.Cols = 2
    mfgLijst.Rows = 2
    
    mfgLijst.TextMatrix(0, 0) = "Btw Nummer"
    mfgLijst.TextMatrix(0, 1) = partnerIs
        
    mfgLijst.ColWidth(0) = 1460
    mfgLijst.ColWidth(1) = 3930
    
    Do While Not rsPartners.EOF
        vatList = vatList & Trim(rsPartners("v150")) & Trim(rsPartners("A161")) & vbCrLf
        vatListAll = vatListAll & Trim(rsPartners("A110")) & vbTab & vbTab & Trim(rsPartners("v150")) & Trim(rsPartners("A161")) & vbTab & vbTab & vbTab & Trim(rsPartners("A100")) & vbCrLf
        mfgLijst.AddItem Trim(rsPartners("v150")) & Trim(rsPartners("A161")) & vbTab & Trim(rsPartners("A100")) & vbTab & Trim(rsPartners("A110"))
        rsPartners.MoveNext
    Loop
    Screen.MousePointer = vbNormal
    
    Me.Height = 7200
    mfgLijst.Visible = True
    Me.cbCopyToClipBoard.Enabled = True

End Sub

Private Sub cbCheckCompanyNumber_Click()

    If Internet_IsAvailable() Then
        r = ShellExecute(0, "open", "https://kbopub.economie.fgov.be/kbopub/zoeknummerform.html?nummer=" + Me.tbCompanyNumber + "&actionLu=Zoek", 0, 0, 1)
    Else
        MsgBox "Geen internet verbinding.", vbInformation
    End If
    
End Sub

Private Sub cbCheckPeppolRegistration_Click()
    
    If Internet_IsAvailable() Then
    Else
        MsgBox "Geen internet verbinding.", vbInformation
        Exit Sub
    End If

    Dim http As Object
    Dim xmlDoc As Object
    Dim docNodes As Object
    Dim docNode As Object
    Dim entityNode As Object
    
    Dim url As String
    Dim responseText As String

    url = "https://directory.peppol.eu/search/1.0/json?q=iso6523-actorid-upis:" + Me.tbPeppolID
    On Local Error Resume Next
    Err = 0

    Screen.MousePointer = vbHourglass
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.Send

    If http.Status = 200 Then
        Screen.MousePointer = vbNormal
        responseText = http.responseText
        
        Me.TextBoxSupportedDocuments = responseText
        MsgBox Me.TextBoxSupportedDocuments
    
        'Set xmlDoc = CreateObject("MSXML2.DOMDocument")
        'xmlDoc.async = False
        'xmlDoc.loadXML responseText
    
        'Set docNodes = xmlDoc.selectNodes("//match/docTypeID")
        'For Each docNode In docNodes
        '    'Debug.Print docNode.Text
        '    MsgBox docNode.Text
        'Next
    
        'Set entityNode = xmlDoc.selectSingleNode("//match/entity")
        'MsgBox "Naam: " & entityNode.selectSingleNode("name").Text
        'MsgBox "Land: " & entityNode.selectSingleNode("countryCode").Text
        'MsgBox "Email: " & entityNode.selectSingleNode("contact").getAttribute("email")
        'Me.TextBoxSupportedDocuments = responseText
    Else
        Screen.MousePointer = vbNormal
        MsgBox "HTTP Error: " & http.Status
        Me.TextBoxSupportedDocuments = ""
    End If
    
    Set http = Nothing
    Set xmlDoc = Nothing
    Set docNodes = Nothing
    Set docNode = Nothing
    Set entityNode = Nothing

End Sub

Private Sub cbCheckVatNumber_Click()
    
    If Internet_IsAvailable() Then
        r = ShellExecute(0, "open", "https://ec.europa.eu/taxation_customs/vies/rest-api/ms/" + Mid(Me.tbVatNumber, 1, 2) + "/vat/" + Mid(Me.tbVatNumber, 3), 0, 0, 1)
    Else
        MsgBox "Geen internet verbinding.", vbInformation
    End If

End Sub

Private Sub cbCopyToClipBoard_Click()
        
    Clipboard.Clear
    Clipboard.SetText vatList
    
    Msg = "Plak de lijst btwnummers in een tekst venster en copy/paste per 25" & vbCrLf
    Msg = Msg & "in toolvenster van bijvoorbeeld https://app.peppolchecker.eu/" & vbCrLf & vbCrLf
    
    If partnerIs = "Leveranciers" Then
        Msg = Msg & "Wat leveranciers betreft, vooral uw aandacht:" & vbCrLf
        Msg = Msg & "Manuele controle van btw nummer en ondernemingsnummer"
    End If
    MsgBox Msg, vbInformation
    
    Dim flHier As Integer
    flHier = FreeFile
    Open LOCATION_COMPANYDATA + "PeppolCheckControle.txt" For Output As flHier
        Print #flHier, vatListAll
    Close flHier
    DoEvents
    Shell "notepad.exe " + LOCATION_COMPANYDATA + "PeppolCheckControle.txt", vbNormalFocus
    flHier = FreeFile
    Open LOCATION_COMPANYDATA + "PeppolEnkelBtwNummers.txt" For Output As flHier
        Print #flHier, vatList
    Close flHier
    DoEvents
    Shell "notepad.exe " + LOCATION_COMPANYDATA + "PeppolEnkelBtwNummers.txt", vbNormalFocus
    
    'r = ShellExecute(0, "open", "https://app.peppolchecker.eu/", 0, 0, 1)
    Unload Me
    
End Sub

Private Sub Form_Load()

    Me.TextBoxSupportedDocuments = ""

End Sub


