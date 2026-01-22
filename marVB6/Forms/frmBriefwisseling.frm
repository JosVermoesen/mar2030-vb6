VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form VrijBericht 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vrije Berichtgeving"
   ClientHeight    =   7650
   ClientLeft      =   -4230
   ClientTop       =   2910
   ClientWidth     =   8730
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   40900
   Icon            =   "frmBriefwisseling.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   13.494
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   15.399
   Begin VB.CheckBox cbPDF 
      Caption         =   "Vpe-Pdf"
      Height          =   315
      Left            =   120
      TabIndex        =   44
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cbFaxBijlage 
      Caption         =   "Bijlage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1680
      Width           =   855
   End
   Begin MSMAPI.MAPIMessages MPIBericht 
      Left            =   4200
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MPISessie 
      Left            =   3600
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton cbBijlage 
      Caption         =   "Bijlage"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cbTogglePrinter 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4920
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   7200
      Width           =   315
   End
   Begin VB.TextBox TxtTekst 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   14
      Left            =   720
      TabIndex        =   39
      Text            =   "v"
      Top             =   1200
      Width           =   2715
   End
   Begin VB.TextBox TxtTekst 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   13
      Left            =   720
      TabIndex        =   38
      Text            =   "v"
      Top             =   1560
      Width           =   2745
   End
   Begin VB.CheckBox cbLokaleAfdruk 
      Caption         =   "Lokale afdruk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3600
      TabIndex        =   37
      Top             =   540
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton cmdFax 
      Caption         =   "F&ax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdMail 
      Caption         =   "eM&ail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Zoeken 
      Caption         =   "&Zoeken"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7560
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   0
      Width           =   1065
   End
   Begin VB.CommandButton Sluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7560
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1065
   End
   Begin VB.CommandButton Wegschrijven 
      Caption         =   "&Afdrukken"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5400
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1065
   End
   Begin VB.CommandButton BerichtInofUit 
      Caption         =   "&Inladen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   7560
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   360
      Width           =   1065
   End
   Begin VB.CommandButton BerichtInofUit 
      Caption         =   "Ops&laan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   7560
      TabIndex        =   30
      Top             =   720
      Width           =   1065
   End
   Begin VB.CommandButton Schoonvegen 
      Caption         =   "Schoo&n"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6480
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1065
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   12
      Left            =   7500
      TabIndex        =   27
      Top             =   1380
      Width           =   1065
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   7500
      TabIndex        =   26
      Top             =   1080
      Width           =   1065
   End
   Begin VB.CheckBox AutoBetreft 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Omschrijving opnemen als 'Betreft' in document"
      Top             =   2100
      Value           =   1  'Checked
      Width           =   225
   End
   Begin VB.CommandButton CmdRtfPrint 
      Caption         =   "Print RTF"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3540
      TabIndex        =   25
      Top             =   60
      Width           =   915
   End
   Begin VB.ComboBox CmbAfdrukModus 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmBriefwisseling.frx":030A
      Left            =   60
      List            =   "frmBriefwisseling.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   60
      Width           =   3465
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Documenttype"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1860
      TabIndex        =   21
      Top             =   360
      Width           =   1635
      Begin VB.OptionButton Briefwisseling 
         Caption         =   "Mailing"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   540
         Width           =   1335
      End
      Begin VB.OptionButton Briefwisseling 
         Caption         =   "Brief"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Partij"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   60
      TabIndex        =   18
      Top             =   360
      Width           =   1605
      Begin VB.OptionButton Partij 
         Caption         =   "Leverancier"
         Height          =   225
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   480
         Width           =   1425
      End
      Begin VB.OptionButton Partij 
         Caption         =   "Klant"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   19
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4680
      TabIndex        =   5
      Top             =   1080
      Width           =   735
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   5430
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   4680
      TabIndex        =   17
      Top             =   1380
      Width           =   735
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   5430
      TabIndex        =   16
      Top             =   1380
      Width           =   2055
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   4680
      TabIndex        =   15
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   7
      Left            =   7590
      TabIndex        =   14
      Top             =   1680
      Width           =   555
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   8
      Left            =   8160
      TabIndex        =   13
      Top             =   1680
      Width           =   405
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   9
      Left            =   4680
      TabIndex        =   12
      Top             =   1980
      Width           =   405
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   10
      Left            =   5100
      TabIndex        =   11
      Top             =   1980
      Width           =   915
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   11
      Left            =   6030
      TabIndex        =   10
      Top             =   1980
      Width           =   2535
   End
   Begin VB.ComboBox Sortering 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4500
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2955
   End
   Begin VB.TextBox TxtTekst 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   2040
      Width           =   3045
   End
   Begin RichTextLib.RichTextBox txtTekst0 
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   8281
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmBriefwisseling.frx":030E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cbDocumentPrinterHier 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   41
      ToolTipText     =   "Dokumenten printer"
      Top             =   7200
      Width           =   3675
   End
   Begin VB.CommandButton cbLayOudPDF 
      Caption         =   "Beheer van VPE Lay-Out"
      Height          =   315
      Left            =   1320
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   7200
      Width           =   2895
   End
   Begin VB.Label lblOorsprong 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   2100
      TabIndex        =   28
      Top             =   3660
      Width           =   45
   End
   Begin VB.Label Label2 
      Caption         =   "&Vrije mededeling"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   3660
      Width           =   1575
   End
   Begin VB.Label VanTot 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4500
      TabIndex        =   8
      Top             =   60
      Width           =   2955
   End
   Begin VB.Label VanTot 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   6060
      TabIndex        =   7
      Top             =   720
      Width           =   1395
   End
   Begin VB.Label VanTot 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   4500
      TabIndex        =   6
      Top             =   720
      Width           =   1545
   End
End
Attribute VB_Name = "VrijBericht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim MaxLijn79 As Integer
Dim KontaktPersoon As Integer

Dim txtFaxIs As String
Dim txtMailIs As String

Dim VsoftVanaf          As String
Dim VsoftTot            As String
Dim BeginXbox           As Integer
Dim BeginYbox           As Integer

Dim FlPartij            As Integer
Dim PartijKode          As String * 1
Dim BriefType           As String * 1

Dim Tel                 As Integer

Dim WProgramma          As String
Dim PNaam               As String
Dim Wdokument           As String
Dim PrintSendKeys       As String
Dim CloseSendKeys       As String
Dim ReedsBestaand As Integer
Dim TekstIsRich As Boolean

Dim rsBrief As ADODB.Recordset
'Dim wApp As Word.Application
Dim mailBijlage As String
Dim faxBijlage As String

'Public g_objFaxDocument As New FAXCOMEXLib.FaxDocument
'Public objFaxVerzender As FaxSender

Private Sub BerichtInOfUit_Click(Index As Integer)
Dim XXX As String

DodaNogess:
On Error Resume Next
XXX = rsBrief.ActiveConnection
If Err Or Index = 0 Then
    Set rsBrief = New ADODB.Recordset
    On Error Resume Next
    Err = 0
    rsBrief.CursorLocation = adUseClient
    Msg = "SELECT * FROM Briefwisseling WHERE A110 = " & _
        "'" + vBibTekst(FlPartij, "#A110 #") + "'"
    
    rsBrief.Open Msg, adntDB, adOpenDynamic, adLockOptimistic
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        Exit Sub
    Else
        'kontrole van de nieuwe velden...
        Err = 0
        On Error Resume Next
        If rsBrief("bstndNaam").Name = "bstndNaam" Then
            'ok
        ElseIf Err = 3265 Then
            rsBrief.Close
            Msg = "ALTER TABLE Briefwisseling ADD COLUMN bstndNaam varchar;"
            Err = 0
            adntDB.Execute Msg
            If Err Then
                MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
            End If
            Msg = "ALTER TABLE Briefwisseling ADD COLUMN typeZending TEXT(5);"
            Err = 0
            adntDB.Execute Msg
            If Err Then
                MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
            End If
            Msg = "ALTER TABLE Briefwisseling ADD COLUMN bstndBLOB OLEobject;"
            Err = 0
            adntDB.Execute Msg
            If Err Then
                MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
            End If
            GoTo DodaNogess
        End If
    End If
End If
        
Select Case Index
    Case 0
        T = 0
        If rsBrief.RecordCount = 0 Then
            MsgBox "Geen briefwisseling te vinden voor deze partij !"
            Exit Sub
        End If
        Unload Xlog
        Load Xlog
        Xlog.Caption = "Bericht inladen voor : " + vBibTekst(FlPartij, "#A100 #")
        Xlog.X.Rows = 1
        Xlog.X.Cols = 2
        Xlog.X.Col = 0
        Xlog.X.Row = 0
        Xlog.X.Text = "Aanmaakdatum"
        Xlog.X.Col = 1
        Xlog.X.Text = "Betreft"
        
        Do While Not rsBrief.EOF
            GoSub VolgendeLijn
            rsBrief.MoveNext
        Loop
        Xlog.WijzigenLijn.Visible = False
        Xlog.Afsluiten.TabStop = False
        Xlog.cbAfbeelding.Visible = False
        Xlog.X.Row = 1
        Xlog.X.Col = 0

        Xlog.X.ColWidth(0) = 1680
        Xlog.X.ColWidth(1) = 4785
        Xlog.X.ColAlignment(0) = flexAlignLeftTop
        Xlog.X.ColAlignment(1) = flexAlignLeftTop
        XLogKey = ""
        Xlog.SSTab1.TabVisible(1) = False
        Xlog.Show 1
        If XLogKey <> "" Then
            ReedsBestaand = True
            rsBrief.MoveFirst
            rsBrief.Move Xlog.X.Row - 1

            On Local Error Resume Next
            VanTot(2).Caption = rsBrief("A110")
            VanTot(0).Caption = rsBrief("v133")
            VanTot(1).Caption = rsBrief("v134")
            lblOorsprong.Caption = rsBrief("v128")
            If rsBrief("v126") = "K" Then
                Partij(0).Value = True
            Else
                Partij(1).Value = True
            End If

            If rsBrief("v127") = "B" Then
                Briefwisseling(0).Value = True
            Else
                Briefwisseling(1).Value = True
            End If
    
            TxtTekst(1).Text = rsBrief("v129")
            If Val(rsBrief("v131")) <= Sortering.ListCount - 1 Then
                Sortering.ListIndex = Val(rsBrief("v131"))
            Else
                Sortering.ListIndex = 0
            End If
            txtTekst0.TextRTF = rsBrief("v132")
            Wegschrijven.Enabled = True
            txtTekst0.Enabled = True
            txtTekst0.SetFocus
        End If
        On Local Error GoTo 0
        Unload Xlog
    
    Case 1
        If Not ReedsBestaand Then
            rsBrief.AddNew
        Else
                  Msg = "Wijzigingen aan bestaand bericht bewaren" + vbCrLf
            Msg = Msg + "Bent U zeker ?"
            KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton1)
            If KtrlBox = vbYes Then
            Else
                Exit Sub
            End If
        End If
        rsBrief("v172") = "16"
        rsBrief("v133") = Left(VanTot(0).Caption, 13)
        rsBrief("v134") = Left(VanTot(1).Caption, 13)
        rsBrief("A110") = VanTot(2).Caption
        rsBrief("v126") = PartijKode
        rsBrief("v127") = BriefType
        rsBrief("v129") = TxtTekst(1).Text
        rsBrief("v130") = Now
        rsBrief("v131") = Sortering.ListIndex
        rsBrief("v132") = txtTekst0.TextRTF
        If ReedsBestaand Then
        Else
            rsBrief("v128") = Now
            If faxBijlage <> "" Then
                FileToBlob rsBrief("bstndBLOB"), faxBijlage
                If Len(faxBijlage) > 255 Then
                    rsBrief("bstndNaam") = Right(faxBijlage, 255)
                Else
                    rsBrief("bstndNaam") = faxBijlage
                End If
                rsBrief("typeZending") = "FAX"
            ElseIf mailBijlage <> "" Then
                FileToBlob rsBrief("bstndBLOB"), mailBijlage
                If Len(mailBijlage) > 255 Then
                    rsBrief("bstndNaam") = Right(mailBijlage, 255)
                Else
                    rsBrief("bstndNaam") = mailBijlage
                End If
                rsBrief("typeZending") = "EMAIL"
            Else
                rsBrief("bstndNaam") = ""
            End If
            'Stop
        End If
        rsBrief.Update
        Schoon
        Zoeken.SetFocus
        
End Select
Exit Sub

VolgendeLijn:
T = T + 1
Select Case FlPartij
    Case TABLE_CUSTOMERS, TABLE_SUPPLIERS
        Xlog.X.AddItem rsBrief("v128") & vbTab & rsBrief("v129")
End Select
Return

End Sub
Private Sub Briefwisseling_Click(Index As Integer)

Select Case Briefwisseling(0).Value
    Case True
        BriefType = "B"
    Case False
        BriefType = "M"
End Select

End Sub


Private Sub cbBijlage_Click()

    If LaadTekst("dnnInstellingen", "Mario") = "" Then
        MsgBox "Nieuwe PC of nog geen instellingen voor PDF Postvak In.  Aanbevolen in te te stellen a.u.b. via submenu DotNetNuke.", vbInformation
        Mim.Teken.InitDir = LOCATION_COMPANYDATA
    Else
        Mim.Teken.InitDir = LaadTekst("dnnInstellingen", "Mario")
    End If
    'Mim.Teken.Filter = "Acrobat bestanden (*.pdf)|*.pdf"
    Mim.Teken.ShowOpen
    If Mim.Teken.filename = "" Then Exit Sub
    mailBijlage = Mim.Teken.filename
    
End Sub

Private Sub cbDocumentPrinterHier_Click()

    dokumentPrinterNr = cbDocumentPrinterHier.ListIndex
    
End Sub

Private Sub cbFaxBijlage_Click()

    Mim.Teken.filename = mailBijlage
    'Mim.Teken.InitDir = LOCATION_COMPANYDATA
    Mim.Teken.ShowOpen
    If Mim.Teken.filename = "" Then Exit Sub
    faxBijlage = Mim.Teken.filename
    
End Sub

Private Sub cbLayOudPDF_Click()

    LayOutpdfDokument.Show
    
End Sub

Private Sub cbPDF_Click()

If cbPDF.Value = vbChecked Then
    Me.cbTogglePrinter.Visible = False
    Me.cbDocumentPrinterHier.Visible = False
Else
    Me.cbTogglePrinter.Visible = True
    Me.cbDocumentPrinterHier.Visible = True
End If

End Sub

Private Sub cbTogglePrinter_Click()

     Me.cbDocumentPrinterHier.Enabled = Not Me.cbDocumentPrinterHier.Enabled
     
End Sub

Private Sub CmbAfdrukModus_Click()

Select Case Left(CmbAfdrukModus.Text, 1)
    Case "0", "1", "3"
        If Me.MPISessie.SessionID <> 0 Then
        Else
            On Local Error Resume Next
            Me.MPISessie.SignOn
            If Err Then
                'Msg = "U kiest om het mailsysteem niet aan te wenden." & vbCr
                'Msg = Msg & "Probeer eventueel later opnieuw E-mail aan te duiden" & vbCr & vbCr
                'Msg = Msg & "Foutmelding van het systeem: " & Err.Description
                'MsgBox Msg, vbExclamation
                SnelHelpPrint Err.Description, BL_LOGGING
                CmbAfdrukModus.ListIndex = 2
            Else
                On Local Error GoTo 0
                Me.MPIBericht.SessionID = Me.MPISessie.SessionID
                SnelHelpPrint "E-mail sessie met succes opgestart. IDkode :" + Format(Me.MPISessie.SessionID), BL_LOGGING
            End If
        End If
    Case "2"
        If Me.MPISessie.SessionID = 0 Then
        Else
            On Local Error Resume Next
            Me.MPISessie.SignOff
            If Err Then
                MsgBox "Fout tijdens afsluiten van een sessie...  Postvak In (InBox) werd voortijdig gesloten."
            Else
                SnelHelpPrint "E-mail sessie werd met succes beëindigd...", BL_LOGGING
            End If
        End If
    Case Else
        MsgBox "stop"
End Select

End Sub


Private Sub cmdFax_Click()
    
    If cmdFax.Enabled = False Then
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
    End If
    
    'As default, create the sender object, and load the default sender information
    Err = 0
    On Error Resume Next
     'Set objFaxVerzender = g_objFaxDocument.Sender
     If Err Then
        MsgBox "Enkel beschikbaar voor Windows XP en hoger", vbInformation
        Screen.MousePointer = vbNormal
        Exit Sub
     End If
     'objFaxVerzender.LoadDefaultSender
     
'Error handling
    'On Error GoTo Error_Handler
    
    'Check if the Aanhangsel checkbox was selected, but no Aanhangsel provided
    'If chkAanhangsel.Value = 1 And txtAanhangselPath.Text = "" Then
    '    MsgBox "Duidt het path voor de bijlage aan en probeer opnieuw.", vbExclamation
    '    Exit Sub
    'End If

    'Remove the previous recipient, if there is one. Otherwise, the fax would be sent to
    'recipients that accumulate in the FaxRecipients collection from each previous use of the form.
    'If g_objFaxDocument.Recipients.Count = 1 Then g_objFaxDocument.Recipients.Remove (1)
    
    'Set no cover page as default
    'g_objFaxDocument.CoverPageType = fcptNONE
        
    'If the checkbox is selected, set cover page type to "local"
    'and set the path and file name to the textbox text
    'If chkVoorBlad.Value = 1 Then
    '    g_objFaxDocument.CoverPageType = fcptLOCAL
    '    g_objFaxDocument.CoverPage = txtCPPath.Text
    'End If
    
    'Set the rest of the fax document properties
    'Noties voorblad
    'g_objFaxDocument.Note = "" 'txtNotities.Text
    'g_objFaxDocument.Subject = "" 'txtOnderwerp.Text
    'g_objFaxDocument.Recipients.Add TxtTekst(13).Text, TxtTekst(3).Text   'txtFaxNummerGeadresseerde.Text, txtNaamGeadresseerde.Text
    'g_objFaxDocument.Sender.Name = objFaxVerzender.Name
    'g_objFaxDocument.Sender.FaxNumber = objFaxVerzender.FaxNumber
    'g_objFaxDocument.Sender.OfficePhone = objFaxVerzender.OfficePhone
    'Include the Aanhangsel if the Aanhangsel checkbox is selected
    
    Dim BestandHier As String
        
    If faxBijlage <> "" Then
        'g_objFaxDocument.Body = faxBijlage
             
        'Submit the document to the server
        'Call g_objFaxDocument.Submit("") '(ServerNameText.Text)
    Else
        'g_objFaxDocument.Body = BestandHier 'txtAanhangselPath.Text
        'g_objFaxDocument.Body = "D:\Mijn Documenten\MODI\test.xls"
     
        'Submit the document to the server
        'Call g_objFaxDocument.Submit("") '(ServerNameText.Text)
        'Kill BestandHier
    End If
    If TxtTekst(13).Text <> txtFaxIs Then
        MsgBox "gewijzigd nummer.  Hier nog routine vraagstelling om al dan niet de partijFiche bij te werken", vbInformation
    End If
    BerichtInOfUit_Click 1
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub cmdMail_Click()

    If cmdMail.Enabled = False Then
        Exit Sub
    ElseIf txtTekst0.Text = "" Then
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
    End If
    
    Dim BestandHier As String
    If Me.MPISessie.SessionID <> 0 Then
    On Local Error GoTo MPIError
    'Compose new message
    Me.MPIBericht.Compose
    
    'Address message
    If CmbAfdrukModus.ListIndex = 0 Then
        Me.MPIBericht.RecipDisplayName = rsMAR(FlPartij)("A100")
        Me.MPIBericht.RecipAddress = "SMTP:" & TxtTekst(14).Text
    Else
        Me.MPIBericht.RecipDisplayName = rsMAR(FlPartij)("A100")
        Me.MPIBericht.RecipAddress = "SMTP:" & TxtTekst(14).Text
    End If
    'Resolve recipient name
    Me.MPIBericht.AddressResolveUI = True '=dialogbox, false = error genereren
    'MpiBericht.ResolveName
        
    'Create the message
    Me.MPIBericht.MsgSubject = TxtTekst(1).Text
    If CmbAfdrukModus.ListIndex = 0 Then
        Me.MPIBericht.MsgNoteText = txtTekst0.Text
    Else
        Me.MPIBericht.MsgSubject = TxtTekst(1).Text
        Me.MPIBericht.MsgNoteText = "Briefwisseling in bijlage als MS Word Document" 'txtTekst0.Text
        'KtrlBox = MsgBox("Verstuurde E-mail aan " & vBibTekst(FlPartij, "#v224 #") & vbCr & vbCr & "Afdruk maken ?", vbQuestion + vbYesNo + vbDefaultButton2)
        Me.MPIBericht.AttachmentPathName = BestandHier
    End If
    If mailBijlage <> "" Then
        Me.MPIBericht.AttachmentPathName = mailBijlage
    End If
    'Send the message
    On Error Resume Next
    Me.MPIBericht.Send True
    If Err = 32001 Then Screen.MousePointer = vbNormal: Exit Sub
    If Err Then MsgBox Error
    If CmbAfdrukModus.ListIndex = 0 Then
    Else
        Kill BestandHier
    End If
    BerichtInOfUit_Click 1
Else
    MsgBox "Uw E-mail systeem is niet geactiveerd.  Contacteer uw systeembeheerder voor bijkomende inlichtingen.", vbExclamation
End If
Screen.MousePointer = vbNormal
Exit Sub

MPIError:
MsgBox Error
Resume Next

End Sub



Private Sub CmdRtfPrint_Click()

'Dim cat As New ADOX.Catalog
'cat.Create adojetprovider & "Data Source=c:\new.mdb"

Mim.Teken.FLAGS = cdlPDReturnDC + cdlPDNoPageNums
If txtTekst0.SelLength = 0 Then
    Mim.Teken.FLAGS = Mim.Teken.FLAGS + cdlPDAllPages
Else
    Mim.Teken.FLAGS = Mim.Teken.FLAGS + cdlPDSelection
End If
Mim.Teken.ShowPrinter
On Error Resume Next
txtTekst0.SelPrint Mim.Teken.hDC

End Sub

Private Sub Form_Load()
Top = 0
Left = 0

cbDocumentPrinterHier.Clear
If Printers.Count = 0 Then MsgBox "Installeer eerst een printerdriver a.u.b. (ook al beschikt U nog over geen printer.  Bij blijvende problemen in netwerk gelieve op te starten met het command '/P=FALSE' en uw netwerkbeheerder te informeren)", vbExclamation: End
For COUNT_TO = 0 To Printers.Count - 1
    Me.cbDocumentPrinterHier.AddItem Printers(COUNT_TO).Port + " " + Printers(COUNT_TO).DeviceName
Next
cbDocumentPrinterHier.ListIndex = dokumentPrinterNr
If String99(READING, 299) = "2" Then
    Me.cbPDF.Value = vbUnchecked
Else
    Me.cbPDF.Value = vbChecked
End If

PAGE_COUNTER = 0
Partij(0).Value = True
txtTekst0.Font = "Courier New"
MaxLijn79 = 90
Briefwisseling(1).Enabled = True
Briefwisseling(0).Value = True

KontaktPersoon = Val(String99(READING, 201))

CmbAfdrukModus.Clear
CmbAfdrukModus.AddItem "0: E-mail via normale tekst"
CmbAfdrukModus.AddItem "1: E-mail via insluiting van Word Sjabloon"
CmbAfdrukModus.AddItem "2: E-mail wordt niet gebruikt op dit ogenblik"
CmbAfdrukModus.ListIndex = 0
CmbAfdrukModus.Enabled = False
Schoon

End Sub

Private Sub Form_Unload(Cancel As Integer)

If Me.MPISessie.SessionID = 0 Then
Else
    On Local Error Resume Next
    Me.MPISessie.SignOff
    If Err Then
        MsgBox "Fout tijdens afsluiten van een sessie...  Postvak In (InBox) werd voortijdig gesloten."
    End If
End If

End Sub

Private Sub Partij_click(Index As Integer)

Select Case Partij(0).Value
    Case True
        PartijKode = "K"
        FlPartij = TABLE_CUSTOMERS
    Case False
        PartijKode = "L"
        FlPartij = TABLE_SUPPLIERS
End Select
BriefType = "B"
Briefwisseling(0).Value = 1

Sortering.Clear
For T = 0 To FL_NUMBEROFINDEXEN(FlPartij)
    Sortering.AddItem Format(T, "00") + ":" + FLINDEX_CAPTION(FlPartij, T)
Next
Sortering.ListIndex = 1

End Sub


Private Sub PrintUserDef(TypeEnTaal As String)
Dim T               As Integer
Dim Tabulatie       As Integer
Dim FlFree          As Integer
Dim Teller          As Integer

ReDim psTekst(50) As String
ReDim psX(50) As Single
ReDim psY(50) As Single
ReDim psFontSize(50) As Single
ReDim psFontName(50) As String
ReDim psFontBold(50) As Integer
ReDim psFontItalic(50) As Integer
ReDim psFontUnderLine(50) As Integer
ReDim psColor(50) As Long
Dim MaxPslokatie        As Integer

Dim BoxTeller As Integer
ReDim Box(10) As Integer
ReDim BeginX(10) As Integer
ReDim BeginY(10) As Integer
ReDim TotX(10) As Integer
ReDim TotY(10) As Integer
Dim PsLokatie As Integer

On Local Error GoTo ErrorLoad

If Dir$(LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + ".Txt") = "" Then
    Beep
    Exit Sub
Else
    FlFree = FreeFile
    Open LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + ".Txt" For Input As FlFree
    Input #FlFree, Dummy
    Input #FlFree, VsoftVanaf
    Input #FlFree, VsoftTot
    For Teller = 0 To 10
        Input #FlFree, BeginX(Teller), BeginY(Teller), TotX(Teller), TotY(Teller), Box(Teller)
    Next
    Teller = 0
    While Not EOF(FlFree)
        Input #FlFree, psTekst(Teller), psX(Teller), psY(Teller), psFontSize(Teller), psFontName(Teller), psFontBold(Teller), psFontItalic(Teller), psFontUnderLine(Teller), psColor(Teller)
        Teller = Teller + 1
    Wend
    MaxPslokatie = Teller
    Close FlFree
End If

BeginXbox = BeginX(0)
BeginYbox = BeginY(0)

PsLokatie = 0
Do While PsLokatie <= MaxPslokatie
    If psTekst(PsLokatie) <> "" Then
        Printer.FontName = psFontName(PsLokatie)
        Printer.FontItalic = psFontItalic(PsLokatie)
        Printer.FontSize = psFontSize(PsLokatie)
        Printer.FontBold = psFontBold(PsLokatie)
        Printer.FontUnderline = psFontUnderLine(PsLokatie)
        Printer.ForeColor = psColor(PsLokatie)
    End If
    Printer.CurrentX = psX(PsLokatie)
    Printer.CurrentY = psY(PsLokatie)
    Printer.Print psTekst(PsLokatie);
    PsLokatie = PsLokatie + 1
Loop
For T = 0 To 10
    If BeginX(T) = 0 Then
    ElseIf Box(T) Then
        Printer.Line (BeginX(T), BeginY(T))-(TotX(T), TotY(T)), , B
    Else
        'printer.Line (BeginX(T), BeginY(T))-(TotX(T), TotY(T))
    End If
Next

Dim FiguurX As Integer
Dim FiguurY As Integer
Dim FiguurName As String
Dim FigBestandsnaam As String
Dim Hoogte As Variant
Dim Breedte As Variant

If Dir$(LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + "G.Txt") = "" Then
Else
    FlFree = FreeFile
    Open LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + "G.Txt" For Input As FlFree
    While Not EOF(FlFree)
        Input #FlFree, FiguurX, FiguurY, FiguurName
        If ScrLeesTekstBestand(Msg, LOCATION_COMPANYDATA & FiguurName & ".mfd") Then
            FigBestandsnaam = Left(Msg, InStr(Msg, vbTab) - 1)
            Mim.imgFiguur.Picture = LoadPicture(FigBestandsnaam)
            Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
            Hoogte = Left(Msg, InStr(Msg, vbTab) - 1)
            Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
            Breedte = Msg
            Printer.PaintPicture Mim.imgFiguur.Picture, FiguurX, FiguurY, Breedte, Hoogte
        Else
            MsgBox "afdrukprobleemke figuurobject: " & FiguurName
        End If
    Wend
    Close FlFree
End If
Exit Sub

ErrorLoad:
MsgBox "Stop tijdens inladen dokumentdefinitie" + vbCrLf + vbCrLf + Error$
Exit Sub
Resume

End Sub

Private Sub Schoon()

VanTot(0).Caption = ""
VanTot(1).Caption = ""
lblOorsprong.Caption = ""
txtTekst0.TextRTF = ""
TxtTekst(1).Text = ""
TxtTekst(13).Text = ""
TxtTekst(14).Text = ""
mailBijlage = ""
faxBijlage = ""

Me.cbBijlage.Enabled = False
Wegschrijven.Enabled = False
txtTekst0.Enabled = False
BerichtInofUit(1).Enabled = False
BerichtInofUit(0).Enabled = True
Zoeken.Enabled = True
TLB_RECORD(TABLE_VARIOUS) = ""

Partij_click 0

End Sub

Private Sub SchoonVegen_Click()

Schoon

End Sub

Private Sub sluiten_Click()

If txtTekst0.Text <> "" Then
          Msg = "Huidige inbreng - al dan niet - reeds weggeschreven.  Venster sluiten." + vbCrLf + vbCrLf
    Msg = Msg + "Bent U zeker ?"
    Ktrl = MsgBox(Msg, 292, "Vrije berichtgeving verlaten")
    If Ktrl = 6 Then
    Else
        Exit Sub
    End If
End If
Unload Me

End Sub



Private Sub TxtTekst_Change(Index As Integer)

'If Index <> 0 Then Exit Sub
'If Not WegSchrijven.Enabled Then
'    WegSchrijven.Enabled = True
'    BerichtInOfUit(1).Enabled = True
'End If
Select Case Index
    Case 13
        If TxtTekst(13).Text = "" Then
            cmdFax.Enabled = False
        Else
            cmdFax.Enabled = True
        End If
        
    Case 14
        If TxtTekst(14).Text = "" Then
            cmdMail.Enabled = False
            Me.cbBijlage.Enabled = False
        Else
            cmdMail.Enabled = True
            Me.cbBijlage.Enabled = True
        End If
End Select


End Sub

Private Sub TxtTekst_GotFocus(Index As Integer)

'If Index <> 0 Then Exit Sub
'Zoeken.Default = False
'TxtTekst(0).FontName = "Courier New"

End Sub

Private Sub TxtTekst_LostFocus(Index As Integer)

'If Index <> 0 Then Exit Sub
Zoeken.Default = True

End Sub

Private Sub TxtTekst0_Change()

If Not Wegschrijven.Enabled Then
    Wegschrijven.Enabled = True
    BerichtInofUit(1).Enabled = True
End If
CmdRtfPrint.Enabled = True

End Sub

Private Sub TxtTekst0_GotFocus()

Zoeken.Default = False
txtTekst0.Font = "Courier New"

End Sub

Private Sub TxtTekst0_LostFocus()

Zoeken.Default = True

End Sub

Private Sub TxtTekst0_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
    Mim.Teken.Filter = "Rich Text Format bestanden|*.rtf"
    Mim.Teken.ShowOpen
    On Error Resume Next
    txtTekst0.LoadFile Mim.Teken.filename, rtfRTF
    If Err Then MsgBox Error
End If


End Sub

Private Sub Wegschrijven_Click()
Dim TelLijn         As Long
Dim positieCRLF     As Long
Dim SpatieLijn      As Long

Dim Tekst           As String
Dim TekstLijn       As String
Dim MijKeuze        As String

If txtTekst0.Text = "" Then
    Exit Sub
ElseIf Me.cbPDF.Value = vbChecked Then
    MsgBox "Stop nog in ontwikkeling"
    Exit Sub
Else
    TelLijn = 0
    Tekst = txtTekst0.Text
    Do While Tekst <> ""
        positieCRLF = InStr(Tekst, vbCrLf)
        If positieCRLF Then
            TekstLijn = Left(Tekst, positieCRLF - 1)
            Tekst = Right(Tekst, Len(Tekst) - positieCRLF - 1)
        Else
            TekstLijn = Tekst
            Tekst = ""
        End If
        TelLijn = TelLijn + 1
        If Len(TekstLijn) > MaxLijn79 Then
            MsgBox "Lijn " + Str$(TelLijn) + " bestaat uit " + Str$(Len(TekstLijn)) + " Tekens. (max." + Str$(MaxLijn79) + " per lijn toegelaten)"
            Exit Sub
        End If
    Loop
End If
If Me.cbPDF.Value = vbChecked Then
    With Mim.Report
        .OpenDoc
        .Author = Trim(String99(READING, 46))
        .GUILanguage = 3 'Nederlands
        .Title = "dnnBriefwisseling"
    End With
    pdfDrukAf
    Mim.Report.WriteDoc (PROGRAM_LOCATION & Format(Now, "YYYYMMDDHHMMSS") & "-dnnBriefwisseling.pdf")
    'If Me.cbPdfExport.Value = vbChecked Then
    '    MsgBox "exporteren naar DNN folder of andere nog te definiëren"
    '    Mim.Report.CloseDoc
    'Else
         Mim.Report.Preview
    'Endif
Else
    DrukAf
End If




On Error Resume Next
VrijBericht.SetFocus

End Sub


Private Sub Zoeken_Click()

Dim JumpTeller As Integer
Dim Taal As String

JumpTeller = -1

Jumper:
JumpTeller = JumpTeller + 1
aIndex = Sortering.ListIndex
SharedFl = FlPartij
GridText = ""
SqlSearch.Show 1
VrijBericht.SetFocus
On Local Error Resume Next
If Trim$(rsMAR(FlPartij)("v224")) <> "" Then   'email
    txtMailIs = rsMAR(FlPartij)("v224")
    TxtTekst(14).Text = txtMailIs
Else
    txtMailIs = ""
    TxtTekst(14).Text = ""
End If
If Trim$(rsMAR(FlPartij)("vs02")) <> "" Then   'fax
    txtFaxIs = rsMAR(FlPartij)("vs02")
    TxtTekst(13).Text = txtFaxIs
Else
    txtFaxIs = ""
    TxtTekst(13).Text = ""
End If

If Ktrl Then
    BerichtInofUit(1).Enabled = False
    Exit Sub
Else
    Msg = vBibTekst(FlPartij, "#A100 #") + vbCrLf
    Msg = Msg + vBibTekst(FlPartij, "#A125 #") + vbCrLf
    Msg = Msg + vBibTekst(FlPartij, "#A104 #") + vBibTekst(FlPartij, "#A105 #") + vBibTekst(FlPartij, "#A106 #") + vbCrLf
    Msg = Msg + vBibTekst(FlPartij, "#A109 #") + " " + vBibTekst(FlPartij, "#A107 #") + " " + vBibTekst(FlPartij, "#A108 #")
    
    txtTekst0.Enabled = True
    VanTot(JumpTeller).Caption = FVT(FlPartij, Sortering.ListIndex)
    VanTot(2).Caption = vBibTekst(FlPartij, "#A110 #")
    If Briefwisseling(0).Value Then
        VanTot(1).Caption = VanTot(0).Caption
    Else
        If JumpTeller Then
        Else
            GoTo Jumper
        End If
    End If
    Taal = vBibTekst(FlPartij, "#A10C #")
    TxtTekst(2) = Mid(fmarBoxText("003", Taal, vBibTekst(FlPartij, "#A102 #")), 4, 10)
    TxtTekst(3) = vBibTekst(FlPartij, "#A100 #")
    TxtTekst(0) = vBibTekst(FlPartij, "#A101 #")
    If Val(vBibTekst(FlPartij, "#A102 #")) = 0 Then
        TxtTekst(2) = ""
    End If
    TxtTekst(4) = Mid(fmarBoxText("003", Taal, vBibTekst(FlPartij, "#vs01 #")), 4, 10)
    TxtTekst(5) = vBibTekst(FlPartij, "#A125 #")
    TxtTekst(12) = vBibTekst(FlPartij, "#A127 #")
    If KontaktPersoon = 1 Then
        If Val(vBibTekst(FlPartij, "#vs01 #")) = 0 Then
            TxtTekst(4) = ""
        End If
    Else
        TxtTekst(4) = ""
        TxtTekst(12) = ""
    End If
    TxtTekst(6) = vBibTekst(FlPartij, "#A104 #")
    TxtTekst(7) = vBibTekst(FlPartij, "#A105 #")
    TxtTekst(8) = vBibTekst(FlPartij, "#A106 #")
    TxtTekst(9) = vBibTekst(FlPartij, "#A109 #")
    TxtTekst(10) = vBibTekst(FlPartij, "#A107 #")
    TxtTekst(11) = vBibTekst(FlPartij, "#A108 #")
    TxtTekst(1).SetFocus
    ReedsBestaand = False
End If

End Sub

Private Sub pdfDrukAf()
Dim X As Integer
Dim T As Integer
Dim BetreftTXT As String
Dim AanspreekTitel As String
Dim Tekst As String
Dim TabIs As Integer
Dim Pagina As Integer
ReDim rSip(5) As String
Dim Taal As String
Dim TabBold As Single
Dim XX As Integer

MousePointer = vbHourglass
Mim.Enabled = False
XDoEvents = DoEvents()
Select Case Briefwisseling(0).Value
    Case True
        bGet FlPartij, 0, vSet((VanTot(2).Caption), 12)
        If Ktrl Then
            MsgBox "Partij '" & Trim$((VanTot(2).Caption)) & "' niet gevonden.  Gelieve klant of leveranciers nogmaals te SELECTeren en opnieuw proberen afdrukken a.u.b."
        Else
            RecordToVeld FlPartij
            GoSub PrintHetPDFdokument
        End If
    Case False
        bGetOrGreater FlPartij, (Sortering.ListIndex), vSet((VanTot(0).Caption), FLINDEX_LEN(FlPartij, (Sortering.ListIndex)))
        If Ktrl Or KEY_BUF(FlPartij) > vSet((VanTot(1).Caption), FLINDEX_LEN(FlPartij, (Sortering.ListIndex))) Then
        Else
            RecordToVeld FlPartij
            GoSub PrintHetPDFdokument
            Do
                bNext FlPartij
                If Ktrl Or KEY_BUF(FlPartij) > vSet((VanTot(1).Caption), FLINDEX_LEN(FlPartij, (Sortering.ListIndex))) Then
                    Exit Do
                Else
                    RecordToVeld FlPartij
                    GoSub PrintHetPDFdokument
                End If
            Loop
        End If
    Case Else
        MsgBox "Partij '" & Trim$((VanTot(2).Caption)) & "' niet gevonden.  Gelieve klant of leveranciers nogmaals te SELECTeren en opnieuw proberen afdrukken a.u.b."
End Select

Printer.EndDoc
Mim.Enabled = True
MousePointer = 0
Exit Sub

PrintHetPDFdokument:
    Tekst = txtTekst0.Text
    TabIs = 10
    GoSub KopBalk
    If DDEFlag <> 1 Then
        GoSub DetailLijnen
        GoSub OnderKant
        Printer.NewPage
    End If
Return

KopBalk:
Pagina = Pagina + 1
Taal = vBibTekst(FlPartij, "#A10C #")
TxtTekst(2) = Mid(fmarBoxText("003", Taal, vBibTekst(FlPartij, "#A102 #")), 4, 10)
TxtTekst(3) = vBibTekst(FlPartij, "#A100 #")
TxtTekst(0) = vBibTekst(FlPartij, "#A101 #")
If Val(vBibTekst(FlPartij, "#A102 #")) = 0 Then
    rSip(0) = TxtTekst(3)
    TxtTekst(2) = ""
Else
    rSip(0) = Trim$(TxtTekst(2)) + " " _
            + Trim$(TxtTekst(3)) + " " _
            + Trim$(TxtTekst(0))
End If
TxtTekst(4) = Mid(fmarBoxText("003", Taal, vBibTekst(FlPartij, "#vs01 #")), 4, 10)
TxtTekst(5) = vBibTekst(FlPartij, "#A125 #")
TxtTekst(12) = vBibTekst(FlPartij, "#A127 #")
If KontaktPersoon = 1 Then
    If Val(vBibTekst(FlPartij, "#vs01 #")) = 0 Then
        rSip(1) = TxtTekst(5)
        TxtTekst(4) = ""
    Else
        rSip(1) = Trim$(TxtTekst(4)) + " " _
                + Trim$(TxtTekst(5)) + " " _
                + Trim$(TxtTekst(12))
    End If
Else
    rSip(1) = ""
    TxtTekst(4) = ""
    TxtTekst(5) = ""
End If
TxtTekst(6) = vBibTekst(FlPartij, "#A104 #")
TxtTekst(7) = vBibTekst(FlPartij, "#A105 #")
TxtTekst(8) = vBibTekst(FlPartij, "#A106 #")
rSip(2) = Trim$(TxtTekst(6)) + " " _
        + Trim$(TxtTekst(7)) + " " _
        + Trim$(TxtTekst(8))

TxtTekst(9) = vBibTekst(FlPartij, "#A109 #")
TxtTekst(10) = vBibTekst(FlPartij, "#A107 #")
TxtTekst(11) = vBibTekst(FlPartij, "#A108 #")
rSip(4) = TxtTekst(9) + " " _
        + TxtTekst(10) + " " _
        + TxtTekst(11)

PrintUserDef "1" + Taal + "4"
Printer.CurrentX = BeginXbox + 350
Printer.CurrentY = BeginYbox + 350
Printer.FontSize = 10
Printer.FontBold = True
For tSip = 0 To 4
    Printer.Print UCase(rSip(tSip)); vbCrLf;
    Printer.CurrentX = BeginXbox + 350
Next

Printer.CurrentY = Val(VsoftVanaf)
Printer.CurrentX = 0
Printer.FontSize = 10
Printer.FontBold = False

Printer.Print Tab(TabIs);
TabBold = Printer.CurrentX

If lblOorsprong.Caption = "" Then
    Printer.Print Tab(60); Format(Now, "DDDD dd mmmm yyyy hh:mm"); vbCrLf; vbCrLf;
Else
    Printer.Print Tab(60); Format(Now, "DDDD dd mmmm yyyy hh:mm"); vbCrLf;
    Printer.Print Tab(60); "(kopij van " + lblOorsprong + ")"; vbCrLf;
End If
Printer.Print vbCrLf;

On Local Error Resume Next
Printer.FontName = txtTekst0.Font

On Local Error GoTo 0

If AutoBetreft.Value = False Then Return

Select Case Partij(0).Value
    Case True
        Select Case Taal
            Case "1"
                BetreftTXT = "Concerne : "
                If vBibTekst(FlPartij, "#A125 #") = "" Then
                    AanspreekTitel = "Cher Cliènt,"
                Else
                    AanspreekTitel = "Cher Cliènts,"
                End If
            Case Else
                BetreftTXT = "Betreft  : "
                If vBibTekst(FlPartij, "#A125 #") = "" Then
                    AanspreekTitel = "Beste Klant,"
                Else
                    AanspreekTitel = "Beste Klanten,"
                End If
        End Select
        Printer.Print Tab(TabIs); AanspreekTitel; vbCrLf; vbCrLf;
    Case Else
        Printer.Print vbCrLf; vbCrLf;
End Select

Printer.FontBold = True
'Printer.CurrentX = TabBold
Printer.Print Tab(TabIs); BetreftTXT; TxtTekst(1).Text; vbCrLf; vbCrLf;
'If Printer.DeviceName = "Fax" Then
    Printer.FontBold = True
'Else
'    Printer.FontBold = False
'End If
Return

DetailLijnen:
Do While Tekst <> ""
    positieCRLF = InStr(Tekst, vbCrLf)
    If positieCRLF Then
        Printer.Print Tab(TabIs); Left(Tekst, positieCRLF - 1); vbCrLf;
        Tekst = Right(Tekst, Len(Tekst) - positieCRLF - 1)
    Else
        Printer.Print Tab(TabIs); Tekst; vbCrLf;
        Tekst = ""
    End If
    If Printer.CurrentY >= Val(VsoftTot) - 100 Then
        Printer.CurrentY = Val(VsoftTot)
        Printer.Print vbCrLf; Tab(TabIs + 60); ".../..."; vbCrLf;
        Printer.NewPage
        Printer.CurrentY = 600
        'goSub KopBalk
    End If
Loop
Return

OnderKant:
Return

PrintHandler:
MsgBox "Stopkode extern toestel." + vbCrLf + vbCrLf + "Kontroleer, of rapporteer aan R&&Vsoft.", 0, "Onverwachte situatie"
Resume

FileHandler:
MsgBox "Onverwachte fout bij openen briefhoofdbestand." + vbCrLf + vbCrLf + "Stopkode : " + Format(Err) + vbCrLf + vbCrLf + "Rapporteer aan R&&Vsoft en/of kontroleer uw setup."
Ktrl = 100
Exit Sub

End Sub


Private Sub DrukAf()
Dim X As Integer
Dim T As Integer
Dim BetreftTXT As String
Dim AanspreekTitel As String
Dim Tekst As String
Dim TabIs As Integer
Dim Pagina As Integer
ReDim rSip(5) As String
Dim Taal As String
Dim TabBold As Single
Dim XX As Integer

If DDEFlag <> 1 Then
    Set Printer = Printers(dokumentPrinterNr)
    On Error Resume Next
    Printer.PaperBin = LaadTekst(App.Title, "dokumentPRINTER")
    If Printer.Orientation = vbPRORLandscape Then
        Printer.Orientation = vbPRORPortrait
        DoEvents
    End If
End If

MousePointer = vbHourglass
Mim.Enabled = False
XDoEvents = DoEvents()
Select Case Briefwisseling(0).Value
    Case True
        bGet FlPartij, 0, vSet((VanTot(2).Caption), 12)
        If Ktrl Then
            MsgBox "Partij '" & Trim$((VanTot(2).Caption)) & "' niet gevonden.  Gelieve klant of leveranciers nogmaals te SELECTeren en opnieuw proberen afdrukken a.u.b."
        Else
            RecordToVeld FlPartij
            GoSub PrintHetdokument
        End If
    Case False
        bGetOrGreater FlPartij, (Sortering.ListIndex), vSet((VanTot(0).Caption), FLINDEX_LEN(FlPartij, (Sortering.ListIndex)))
        If Ktrl Or KEY_BUF(FlPartij) > vSet((VanTot(1).Caption), FLINDEX_LEN(FlPartij, (Sortering.ListIndex))) Then
        Else
            RecordToVeld FlPartij
            GoSub PrintHetdokument
            Do
                bNext FlPartij
                If Ktrl Or KEY_BUF(FlPartij) > vSet((VanTot(1).Caption), FLINDEX_LEN(FlPartij, (Sortering.ListIndex))) Then
                    Exit Do
                Else
                    RecordToVeld FlPartij
                    GoSub PrintHetdokument
                End If
            Loop
        End If
    Case Else
        MsgBox "Partij '" & Trim$((VanTot(2).Caption)) & "' niet gevonden.  Gelieve klant of leveranciers nogmaals te SELECTeren en opnieuw proberen afdrukken a.u.b."
End Select
Printer.EndDoc
Mim.Enabled = True
MousePointer = 0
Exit Sub

PrintHetdokument:
    Tekst = txtTekst0.Text
    TabIs = 10
    GoSub KopBalk
    If DDEFlag <> 1 Then
        GoSub DetailLijnen
        GoSub OnderKant
        Printer.NewPage
    End If
Return

KopBalk:
Pagina = Pagina + 1
Taal = vBibTekst(FlPartij, "#A10C #")
TxtTekst(2) = Mid(fmarBoxText("003", Taal, vBibTekst(FlPartij, "#A102 #")), 4, 10)
TxtTekst(3) = vBibTekst(FlPartij, "#A100 #")
TxtTekst(0) = vBibTekst(FlPartij, "#A101 #")
If Val(vBibTekst(FlPartij, "#A102 #")) = 0 Then
    rSip(0) = TxtTekst(3)
    TxtTekst(2) = ""
Else
    rSip(0) = Trim$(TxtTekst(2)) + " " _
            + Trim$(TxtTekst(3)) + " " _
            + Trim$(TxtTekst(0))
End If
TxtTekst(4) = Mid(fmarBoxText("003", Taal, vBibTekst(FlPartij, "#vs01 #")), 4, 10)
TxtTekst(5) = vBibTekst(FlPartij, "#A125 #")
TxtTekst(12) = vBibTekst(FlPartij, "#A127 #")
If KontaktPersoon = 1 Then
    If Val(vBibTekst(FlPartij, "#vs01 #")) = 0 Then
        rSip(1) = TxtTekst(5)
        TxtTekst(4) = ""
    Else
        rSip(1) = Trim$(TxtTekst(4)) + " " _
                + Trim$(TxtTekst(5)) + " " _
                + Trim$(TxtTekst(12))
    End If
Else
    rSip(1) = ""
    TxtTekst(4) = ""
    TxtTekst(5) = ""
End If
TxtTekst(6) = vBibTekst(FlPartij, "#A104 #")
TxtTekst(7) = vBibTekst(FlPartij, "#A105 #")
TxtTekst(8) = vBibTekst(FlPartij, "#A106 #")
rSip(2) = Trim$(TxtTekst(6)) + " " _
        + Trim$(TxtTekst(7)) + " " _
        + Trim$(TxtTekst(8))

TxtTekst(9) = vBibTekst(FlPartij, "#A109 #")
TxtTekst(10) = vBibTekst(FlPartij, "#A107 #")
TxtTekst(11) = vBibTekst(FlPartij, "#A108 #")
rSip(4) = TxtTekst(9) + " " _
        + TxtTekst(10) + " " _
        + TxtTekst(11)

PrintUserDef "1" + Taal + "4"
Printer.CurrentX = BeginXbox + 350
Printer.CurrentY = BeginYbox + 350
Printer.FontSize = 10
Printer.FontBold = True
For tSip = 0 To 4
    Printer.Print UCase(rSip(tSip)); vbCrLf;
    Printer.CurrentX = BeginXbox + 350
Next

Printer.CurrentY = Val(VsoftVanaf)
Printer.CurrentX = 0
Printer.FontSize = 10
Printer.FontBold = False

Printer.Print Tab(TabIs);
TabBold = Printer.CurrentX

If lblOorsprong.Caption = "" Then
    Printer.Print Tab(60); Format(Now, "DDDD dd mmmm yyyy hh:mm"); vbCrLf; vbCrLf;
Else
    Printer.Print Tab(60); Format(Now, "DDDD dd mmmm yyyy hh:mm"); vbCrLf;
    Printer.Print Tab(60); "(kopij van " + lblOorsprong + ")"; vbCrLf;
End If
Printer.Print vbCrLf;

On Local Error Resume Next
Printer.FontName = txtTekst0.Font

On Local Error GoTo 0

If AutoBetreft.Value = False Then Return

Select Case Partij(0).Value
    Case True
        Select Case Taal
            Case "1"
                BetreftTXT = "Concerne : "
                If vBibTekst(FlPartij, "#A125 #") = "" Then
                    AanspreekTitel = "Cher Cliènt,"
                Else
                    AanspreekTitel = "Cher Cliènts,"
                End If
            Case Else
                BetreftTXT = "Betreft  : "
                If vBibTekst(FlPartij, "#A125 #") = "" Then
                    AanspreekTitel = "Beste Klant,"
                Else
                    AanspreekTitel = "Beste Klanten,"
                End If
        End Select
        Printer.Print Tab(TabIs); AanspreekTitel; vbCrLf; vbCrLf;
    Case Else
        On Local Error Resume Next
        Printer.Print vbCrLf; vbCrLf;
End Select

Printer.FontBold = True
'Printer.CurrentX = TabBold
Printer.Print Tab(TabIs); BetreftTXT; TxtTekst(1).Text; vbCrLf; vbCrLf;
'If Printer.DeviceName = "Fax" Then
    Printer.FontBold = True
'Else
'    Printer.FontBold = False
'End If
Return

DetailLijnen:
If InStr(Tekst, "@<") Then
    GoSub InsertVars
End If

Do While Tekst <> ""
    positieCRLF = InStr(Tekst, vbCrLf)
    If positieCRLF Then
        Printer.Print Tab(TabIs); Left(Tekst, positieCRLF - 1); vbCrLf;
        Tekst = Right(Tekst, Len(Tekst) - positieCRLF - 1)
    Else
        Printer.Print Tab(TabIs); Tekst; vbCrLf;
        Tekst = ""
    End If
    If Printer.CurrentY >= Val(VsoftTot) - 100 Then
        Printer.CurrentY = Val(VsoftTot)
        Printer.Print vbCrLf; Tab(TabIs + 60); ".../..."; vbCrLf;
        Printer.NewPage
        Printer.CurrentY = 600
        'goSub KopBalk
    End If
Loop
Return

InsertVars:
Dim pVARbegin As Integer
Dim pVARend As Integer
Dim pVARcode As String
Dim pVARhier As String

Do While InStr(Tekst, "@<")
    pVARbegin = InStr(Tekst, "@<")
    pVARend = InStr(Tekst, ">@")
    'checken of het ongeveer klopt
    If pVARend - pVARbegin <> 7 Then
        MsgBox ("Onlogische variabelpositie ter hoogte: " & Mid(Tekst, pVARbegin))
        Exit Do
    Else
        'ok
        pVARcode = Mid(Tekst, pVARbegin, 9)
        On Error Resume Next
        Err = 0
        If IsNull(rsMAR(FlPartij)(Trim(Mid(pVARcode, 3, 5)))) Then
            pVARhier = ""
            If Err Then
                MsgBox Err.Description & vbCrLf & vbCrLf & Mid(pVARcode, 3, 5)
            End If
        Else
            pVARhier = rsMAR(FlPartij)(Trim(Mid(pVARcode, 3, 5)))
        End If
        Tekst = Replace(Tekst, pVARcode, pVARhier)
    End If
Loop
Return


OnderKant:
Return

PrintHandler:
MsgBox "Stopkode extern toestel." + vbCrLf + vbCrLf + "Kontroleer, of rapporteer aan R&&Vsoft.", 0, "Onverwachte situatie"
Resume

FileHandler:
MsgBox "Onverwachte fout bij openen briefhoofdbestand." + vbCrLf + vbCrLf + "Stopkode : " + Format(Err) + vbCrLf + vbCrLf + "Rapporteer aan R&&Vsoft en/of kontroleer uw setup."
Ktrl = 100
Exit Sub

End Sub



' Copy a BLOB field's contents to a binary file.
Function BlobToFile(fld As ADODB.Field, filename As String, _
    Optional ChunkSize As Long = 8192)
    Dim fnum As Integer, bytesLeft As Long, bytes As Long
    Dim tmp() As Byte
    
    ' Raise an error if the field doesn't support GetChunk.
    If (fld.Attributes And adFldLong) = 0 Then
        Err.Raise 1001, , "Field doesn't support the GetChunk method."
    End If
    ' Open the file;, delete it firstoverwrite it if necessary.' Delete the
    ' file if it exists already, then create a new one.
    If Dir$(filename) <> "" Then Kill filename
    
    fnum = FreeFile
    Open filename For Binary As fnum
    ' Read the field's contents, and write it the data to the file.
    bytesLeft = fld.ActualSize
    
    Do While bytesLeft
        bytes = bytesLeft
        If bytes > ChunkSize Then bytes = ChunkSize
        tmp = fld.GetChunk(bytes)
        Put #fnum, , tmp
        bytesLeft = bytesLeft - bytes
    Loop
    Close #fnum
    
End Function


' Copy a file's contents into a BLOB field.
Function FileToBlob(fld As ADODB.Field, filename As String, _
    Optional ChunkSize As Long = 8192)
    Dim fnum As Integer, bytesLeft As Long, bytes As Long
    Dim tmp() As Byte
    
    ' Raise an error if the field doesn't support GetChunk.
    If (fld.Attributes And adFldLong) = 0 Then
        Err.Raise 1001, , "Field doesn't support the GetChunk method."
    End If
    ' Open the file; raise an error if the file doesn't exist.
    If Dir$(filename) = "" Then Err.Raise 53, , "File not found"
    
    fnum = FreeFile
    Open filename For Binary As fnum
    ' Read the file in chunks, and append data to the field.
    bytesLeft = LOF(fnum)
    Do While bytesLeft
        bytes = bytesLeft
        If bytes > ChunkSize Then bytes = ChunkSize
        ReDim tmp(1 To bytes) As Byte
        Get #1, , tmp
        fld.AppendChunk tmp
        bytesLeft = bytesLeft - bytes
    Loop
    
    Close #fnum
    
End Function


