VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form As2TelebibIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inkomende uitwisseling"
   ClientHeight    =   6105
   ClientLeft      =   3120
   ClientTop       =   2430
   ClientWidth     =   9960
   ControlBox      =   0   'False
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
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6105
   ScaleWidth      =   9960
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cbTestTB2Table 
      Caption         =   "TB2 TT"
      Height          =   375
      Left            =   6840
      TabIndex        =   43
      Top             =   120
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4335
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   7646
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&AS/2"
      TabPicture(0)   =   "frmAS2Telebib2IN.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "RichTextBox1"
      Tab(0).Control(1)=   "lbBerichten"
      Tab(0).Control(2)=   "cbAS2BerichtType"
      Tab(0).Control(3)=   "cmdPManueel"
      Tab(0).Control(4)=   "cmdEManueel"
      Tab(0).Control(5)=   "cbPrintNote"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "&Historiek"
      TabPicture(1)   =   "frmAS2Telebib2IN.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "RichTextBoxHistoriek"
      Tab(1).Control(1)=   "cbAS2BerichtTypeHistoriek"
      Tab(1).Control(2)=   "KeuzeInfo(0)"
      Tab(1).Control(3)=   "lbBerichtenHistoriek"
      Tab(1).Control(4)=   "cbNogEens"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "GBO"
      TabPicture(2)   =   "frmAS2Telebib2IN.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "obVerwerking(1)"
      Tab(2).Control(1)=   "obVerwerking(0)"
      Tab(2).Control(2)=   "oGBO300"
      Tab(2).Control(3)=   "opRekeningUittreksel"
      Tab(2).Control(4)=   "obBorderel"
      Tab(2).Control(5)=   "cbFile"
      Tab(2).Control(6)=   "cbGBO"
      Tab(2).Control(7)=   "RichTextBox2"
      Tab(2).Control(8)=   "Label1(4)"
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "MailBox"
      TabPicture(3)   =   "frmAS2Telebib2IN.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label2"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label3"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "rtbMail"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "cbMailBerichten"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "cbFetch"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtMailBox"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "cbMailManueel"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "tbTelRetours"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "AS/2 no Outlook"
      TabPicture(4)   =   "frmAS2Telebib2IN.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(5)"
      Tab(4).Control(1)=   "RichTextBox3"
      Tab(4).Control(2)=   "aswebcbFile"
      Tab(4).Control(3)=   "cbASWEB"
      Tab(4).ControlCount=   4
      Begin VB.CommandButton cbASWEB 
         Caption         =   "Verwerken"
         Height          =   375
         Left            =   -66720
         TabIndex        =   41
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton aswebcbFile 
         Caption         =   "..."
         Default         =   -1  'True
         Height          =   375
         Left            =   -66000
         TabIndex        =   39
         Top             =   360
         Width           =   495
      End
      Begin VB.TextBox tbTelRetours 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   38
         Text            =   "0"
         Top             =   720
         Width           =   1815
      End
      Begin VB.CommandButton cbMailManueel 
         Caption         =   "Verwerken"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7680
         TabIndex        =   36
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox txtMailBox 
         Height          =   285
         Left            =   5760
         TabIndex        =   34
         Text            =   "rbr@rv.be"
         Top             =   450
         Width           =   1815
      End
      Begin VB.CommandButton cbFetch 
         Caption         =   "Fetch"
         Height          =   495
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   1455
      End
      Begin VB.ComboBox cbMailBerichten 
         Height          =   315
         ItemData        =   "frmAS2Telebib2IN.frx":008C
         Left            =   240
         List            =   "frmAS2Telebib2IN.frx":008E
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1080
         Width           =   9255
      End
      Begin VB.OptionButton obVerwerking 
         Alignment       =   1  'Right Justify
         Caption         =   "Uitwisseling enkel opslaan (verwerking later)"
         Height          =   195
         Index           =   1
         Left            =   -69720
         TabIndex        =   30
         Top             =   3960
         Width           =   4215
      End
      Begin VB.OptionButton obVerwerking 
         Caption         =   "Directe verwerking (synchroon)"
         Height          =   195
         Index           =   0
         Left            =   -74880
         TabIndex        =   29
         Top             =   3960
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton oGBO300 
         Caption         =   "Ruilkaart 300"
         Height          =   435
         Left            =   -71760
         TabIndex        =   28
         ToolTipText     =   "Interessant voor import naam en volledig adres!"
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton opRekeningUittreksel 
         Caption         =   "TB2 Rek. Uittreksel"
         Height          =   435
         Left            =   -73320
         TabIndex        =   27
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton obBorderel 
         Caption         =   "TB2 Borderel"
         Height          =   435
         Left            =   -74760
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.CommandButton cbPrintNote 
         Caption         =   "Print IPM Note"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -73080
         TabIndex        =   25
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cbFile 
         Caption         =   "..."
         Height          =   375
         Left            =   -66000
         TabIndex        =   23
         Top             =   480
         Width           =   495
      End
      Begin VB.CommandButton cbGBO 
         Caption         =   "Verwerken"
         Height          =   375
         Left            =   -66720
         TabIndex        =   22
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton cbNogEens 
         Caption         =   "Verwerken (herhaal)"
         Height          =   495
         Left            =   -67200
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   420
         Width           =   1575
      End
      Begin VB.ComboBox lbBerichtenHistoriek 
         Height          =   315
         ItemData        =   "frmAS2Telebib2IN.frx":0090
         Left            =   -74880
         List            =   "frmAS2Telebib2IN.frx":0092
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1440
         Width           =   9255
      End
      Begin VB.ComboBox KeuzeInfo 
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
         Index           =   0
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   7545
      End
      Begin VB.ComboBox cbAS2BerichtTypeHistoriek 
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
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   9255
      End
      Begin VB.CommandButton cmdEManueel 
         Caption         =   "AS2 Verwerken"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -70440
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton cmdPManueel 
         Caption         =   "AS1 Positioneel"
         Enabled         =   0   'False
         Height          =   375
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cbAS2BerichtType 
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
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   960
         Width           =   9495
      End
      Begin VB.ComboBox lbBerichten 
         Height          =   315
         ItemData        =   "frmAS2Telebib2IN.frx":0094
         Left            =   -74880
         List            =   "frmAS2Telebib2IN.frx":0096
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1320
         Width           =   9495
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   2
         Top             =   1680
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   3625
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmAS2Telebib2IN.frx":0098
      End
      Begin RichTextLib.RichTextBox RichTextBoxHistoriek 
         Height          =   1995
         Left            =   -74880
         TabIndex        =   6
         Top             =   1800
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3519
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmAS2Telebib2IN.frx":017E
      End
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   21
         Top             =   1440
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4260
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmAS2Telebib2IN.frx":0264
      End
      Begin RichTextLib.RichTextBox rtbMail 
         Height          =   2775
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4895
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmAS2Telebib2IN.frx":033E
      End
      Begin RichTextLib.RichTextBox RichTextBox3 
         Height          =   2415
         Left            =   -74880
         TabIndex        =   42
         Top             =   1320
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   4260
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmAS2Telebib2IN.frx":0418
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Manueel"
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   40
         Top             =   480
         Width           =   8610
      End
      Begin VB.Label Label3 
         Caption         =   "Ontvangen Retour"
         Height          =   255
         Left            =   3960
         TabIndex        =   37
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Te Controleren Mailbox"
         Height          =   255
         Left            =   3600
         TabIndex        =   35
         Top             =   480
         Width           =   2295
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Manueel"
         Height          =   255
         Index           =   4
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   8610
      End
   End
   Begin VB.CommandButton Annuleren 
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
      Height          =   300
      Left            =   5400
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1245
   End
   Begin MSMAPI.MAPIMessages MPIBericht 
      Left            =   7560
      Top             =   720
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
      Left            =   6960
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label lblAantalBerichten 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   3480
      TabIndex        =   18
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label UserInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   675
      Index           =   1
      Left            =   2040
      TabIndex        =   13
      Top             =   390
      Width           =   4635
   End
   Begin VB.Label UserInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   11
      Top             =   30
      Width           =   4635
   End
   Begin VB.Label UserInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   3
      Left            =   2310
      TabIndex        =   15
      Top             =   1320
      Width           =   1140
   End
   Begin VB.Label UserInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1320
      Width           =   1140
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aard van het verzoek"
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   10
      Top             =   30
      Width           =   1935
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Maatschappij"
      Height          =   285
      Index           =   1
      Left            =   750
      TabIndex        =   12
      Top             =   390
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Manueel"
      Height          =   270
      Index           =   3
      Left            =   1320
      TabIndex        =   17
      Top             =   1320
      Width           =   930
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Polissen in Fiche aangepast"
      Height          =   525
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   1530
   End
End
Attribute VB_Name = "As2TelebibIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim pdfY As Double
Dim tellertje As Integer
Dim lijnTeller As Integer
Dim troelala$

Dim MaandVerwerking As String
Dim BoekjaarKontrole As String
Dim DagKwijting  As String
Dim MaandKwijting As String
Dim DatumKwijting As String

Dim Bedrag0 As Double
Dim Bedrag1 As Double
Dim Bedrag2 As Double
Dim Bedrag3 As Double
Dim Bedrag4 As Double
Dim Bedrag5 As Double
Dim Bedrag6 As Double
Dim Bedrag7 As Double
Dim Bedrag8 As Double
Dim Bedrag9 As Double



Dim FlAsIn As Integer
Dim FlAsNet As Integer
Dim FlSchrijfWeg As Integer
Dim A As String
Dim aa As String

Dim KlantNaam1 As String
Dim Klantnaam2 As String
Dim KlantStraat As String
Dim KlantPlaats As String
Dim KlantTelefoon As String
Dim KlantFax As String

Dim DetailPolis1 As String
Dim DetailPolis2 As String

Dim TLBCode(100) As String * 10
Dim TLBTekst(100) As String
Dim TLBType(100) As String
Dim TLBLengte(100) As Integer
Dim TLBPos(100) As Integer

Dim rsTB2 As ADODB.Recordset

Dim gboDIR As String
Dim gboFILE As String
Dim gboCPstr As String

Dim aswebDIR As String
Dim aswebFILE As String
Dim aswebCPstr As String


Dim strA000RB As String


Private Sub Annuleren_Click()

    Unload Me

End Sub

Private Sub aswebcbFile_Click()

On Local Error Resume Next
    Mim.Teken.InitDir = aswebDIR
    Mim.Teken.filename = aswebDIR & aswebFILE
    Mim.Teken.ShowOpen
    If Err Then
        Mim.Teken.filename = " "
    End If
    If Mim.Teken.filename = "" Then
        Exit Sub
    ElseIf Mim.Teken.filename <> aswebDIR Then
        aswebDIR = UCase(Mim.Teken.filename)
        aswebDIR = Left(aswebDIR, InStr(aswebDIR, aswebFILE) - 1)
        BeWaarTekst App.Title, "aswebDIR", aswebDIR
        Label1(5).Caption = aswebDIR & aswebFILE
        Me.cbASWEB.Enabled = True
        'CURRAC= rekeninguittreksel
        'PRENOT= termijnborderel
        If ScrLeesBestandAlleTekst(aswebCPstr, aswebDIR & aswebFILE) Then
            Me.RichTextBox3.Text = aswebCPstr
        Else
            Me.RichTextBox3.Text = ""
        End If
    End If


End Sub

Private Sub cbAS2BerichtType_Click()

    VernieuwMAPIBerichten

End Sub


Private Sub cbAS2BerichtTypeHistoriek_Click()

   KeuzeInfo_Click 0
    
End Sub

Private Sub cbASWEB_Click()

    If Trim$(Me.RichTextBox3.Text) = "" Then Exit Sub
    Dim atXGHarray() As String
    
    Do While InStr(Me.RichTextBox3.Text, "XGT+01'")
        MsgBox "RTB check for XGT+1"
        Me.RichTextBox3.Text = Replace(Me.RichTextBox3.Text, "XGT+01'", "XGT+1'")
    Loop
    
    
    atXGHarray = Split(Me.RichTextBox3.Text, "XGT+1'")
    For COUNT_TO = 0 To UBound(atXGHarray()) - 1
        atXGHarray(COUNT_TO) = atXGHarray(COUNT_TO) & "XGT+1'"
    Next
    
    Dim tb2infoHier As String

    With Mim.Report
        .OpenDoc
        .Author = Trim(String99(READING, 46))
        .GUILanguage = 3 'Nederlands
        .Title = "dnnInning4Brokers"
        .FontName = "Courier New"
        .FontSize = 8
        .TextBold = True
        .TextColor = 0 'zwart
        .nTopMargin = 1
        .nLeftMargin = 0.5
        .nRightMargin = 0.5
        .PenSize = 0.01
    End With
    
   
    For COUNT_TO = 0 To UBound(atXGHarray()) - 1
        RichTextBox1.Text = atXGHarray(COUNT_TO)
        
        tb2infoHier = tb2Indent(Me.RichTextBox1.Text)
        pdfY = Mim.Report.VPEPRINT(1, 1, tb2infoHier & vbCrLf)
        Mim.Report.PageBreak
        MsgBox "Start " + Str(COUNT_TO + 1)
        KtrlBox = DoTheGBO()
    Next
    
    Mim.Report.WriteDoc (PROGRAM_LOCATION & Format(Now, "YYYYMMDDHHMMSS") & "-dnnAS1_EDI.pdf")
    Mim.Report.Preview
    '!Mim.Report.CloseDoc
    
ASWEBeinde:
    KtrlBox = MsgBox("Hierna wordt het AS1.EDI bestand verwijderd (alle ontvangen berichten worden inmmers automatisch in het logboek opgenomen.)" & vbCrLf & vbCrLf & "Verwijderen GBO bestand, bent U zeker", vbYesNoCancel + vbQuestion + vbDefaultButton3)
    If KtrlBox = vbYes Then
        Err = 0
        On Local Error Resume Next
        Kill Label1(5).Caption
        If Err Then
            MsgBox Error
        Else
            Me.RichTextBox3.Text = ""
        End If
    ElseIf KtrlBox = vbCancel Then
        GoTo ASWEBeinde
    Else
        'Stop
    End If

End Sub

Private Sub cbFile_Click()

    On Local Error Resume Next
    Mim.Teken.InitDir = gboDIR
    Mim.Teken.filename = gboDIR & gboFILE
    Mim.Teken.ShowOpen
    If Err Then
        Mim.Teken.filename = " "
    End If
    If Mim.Teken.filename = "" Then
        Exit Sub
    ElseIf Mim.Teken.filename <> gboDIR Then
        gboDIR = UCase(Mim.Teken.filename)
        gboDIR = Left(gboDIR, InStr(gboDIR, gboFILE) - 1)
        BeWaarTekst App.Title, "gboDIR", gboDIR
        Label1(4).Caption = gboDIR & gboFILE
        Me.cbGBO.Enabled = True
        'CURRAC= rekeninguittreksel
        'PRENOT= termijnborderel
        If ScrLeesBestandAlleTekst(gboCPstr, gboDIR & gboFILE) Then
            Me.RichTextBox2.Text = gboCPstr
        Else
            Me.RichTextBox2.Text = ""
        End If
    End If

End Sub

Private Sub cbGBO_Click()
    
    If Right(Label1(4).Caption, 10) = "UNIFEP.300" Then
        DoTheGBO300
        Exit Sub
    End If
    

    If Trim$(Me.RichTextBox2.Text) = "" Then Exit Sub
    Dim atXGHarray() As String
    
    atXGHarray = Split(Me.RichTextBox2.Text, "XGT+1'")
    For COUNT_TO = 0 To UBound(atXGHarray()) - 1
        atXGHarray(COUNT_TO) = atXGHarray(COUNT_TO) & "XGT+1'"
    Next
    
    Dim tb2infoHier As String

    With Mim.Report
        .OpenDoc
        .Author = Trim(String99(READING, 46))
        .GUILanguage = 3 'Nederlands
        .Title = "dnnInning4Brokers"
        .FontName = "Courier New"
        .FontSize = 8
        .TextBold = True
        .TextColor = 0 'zwart
        .nTopMargin = 1
        .nLeftMargin = 0.5
        .nRightMargin = 0.5
        .PenSize = 0.01
    End With
     
    
    For COUNT_TO = 0 To UBound(atXGHarray()) - 1
        RichTextBox1.Text = atXGHarray(COUNT_TO)
        tb2infoHier = tb2Indent(Me.RichTextBox1.Text)
        pdfY = Mim.Report.VPEPRINT(1, 1, tb2infoHier & vbCrLf)
        Mim.Report.PageBreak
        KtrlBox = DoTheGBO()
    Next
    
    Mim.Report.WriteDoc (PROGRAM_LOCATION & Format(Now, "YYYYMMDDHHMMSS") & "-dnnAS1_EDI.pdf")
    Mim.Report.Preview
    '!Mim.Report.CloseDoc
    
GBOeinde:
    KtrlBox = MsgBox("Hierna wordt het GBO bestand verwijderd (alle ontvangen berichten worden inmmers automatisch in het logboek opgenomen.)" & vbCrLf & vbCrLf & "Verwijderen GBO bestand, bent U zeker", vbYesNoCancel + vbQuestion + vbDefaultButton3)
    If KtrlBox = vbYes Then
        Err = 0
        On Local Error Resume Next
        Kill Label1(4).Caption
        If Err Then MsgBox Error
        opRekeningUittreksel.Value = Not opRekeningUittreksel.Value
    ElseIf KtrlBox = vbCancel Then
        GoTo GBOeinde
    Else
        'Stop
    End If
    
    
End Sub


Private Sub cbMailBerichten_Click()

    Me.MPIBericht.MsgIndex = Me.cbMailBerichten.ListIndex
    If InStr(Me.MPIBericht.RecipAddress, Me.txtMailBox.Text) Then
        If Me.MPIBericht.AttachmentCount <> 0 Then
            Me.rtbMail.LoadFile MPIBericht.AttachmentPathName
            Me.cbMailManueel.Enabled = True
        Else
            Me.rtbMail.Text = "Er is geen TB2 bijlage?"
            Me.cbMailManueel.Enabled = False
        End If
    Else
        Me.rtbMail.Text = Me.MPIBericht.MsgNoteText
        Me.cbMailManueel.Enabled = False
    End If
    Me.cbMailBerichten.SetFocus
    
End Sub

Private Sub cbNogEens_Click()


    If Trim$(RichTextBoxHistoriek.Text) = "" Then Exit Sub
    Msg = "Bericht nog eens doorsturen"
    If MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then
        RichTextBox1.Text = RichTextBoxHistoriek.Text
        cmdEManueel_Click
    End If
    
End Sub

Private Sub cbPrintNote_Click()

    On Local Error Resume Next
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = True
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print RichTextBox1.Text
    Printer.FontBold = False
    Printer.EndDoc
    GoSub VerwijderPRINTnote
Exit Sub
    
VerwijderPRINTnote:
ProbeerHierNogEens:
KtrlBox = MsgBox("PRINTNOTE Bericht verwijderen.  Bent U zeker" & vbCrLf & "(Druk 'Ja' om te verwijderen, een copij wordt altijd bewaard!)", vbYesNoCancel + vbQuestion + vbDefaultButton3)
If KtrlBox = vbCancel Then
    GoTo ProbeerHierNogEens
ElseIf KtrlBox = vbNo Then
Else
    'TLB_RECORD(TABLE_VARIOUS) = ""
    'vBib TABLE_VARIOUS, vSet("L" + "CO" + Format(XGHarray(2), "0000"), 13), "v004"
    'vBib TABLE_VARIOUS, Format(XGHarray(2), "0000"), "A010"
    'vBib TABLE_VARIOUS, "E1", "v400"
    'vBib TABLE_VARIOUS, vSet("30" + Format(XGHarray(2), "0000") + "E1", 20), "v005"
    'vBib TABLE_VARIOUS, RichTextBox1.Text, "v132"
    'bInsert TABLE_VARIOUS, 1
    'If Ktrl Then
    'Else
        'On Local Error Resume Next
        'MPIBericht.Delete
        'VernieuwMAPIBerichten
End If
Return

End Sub

Private Sub cbTestTB2Table_Click()

    Dim rcTB2 As ADODB.Recordset
    Dim strDocType As String
    Dim strMemoTB2 As String
    Dim strSub As String
    
    Dim searchString As String
    Dim startIndex As Integer
    Dim endIndex As Integer
            
    On Error Resume Next
    Set rcTB2 = New ADODB.Recordset
    
    Err = 0
    Msg = "SELECT * FROM Polissen"
    rcTB2.CursorLocation = adUseClient
    rcTB2.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
    If Err Then MsgBox "stop"
    rcTB2.MoveFirst
    
    Do While Not rcTB2.EOF
        strDocType = RTrim(rcTB2(1))
        If strDocType = strSub Then
            strDocType = strDocType + Trim(Str(tellertje))
            rcTB2(1) = strDocType
        Else
            strSub = strDocType
        End If
        rcTB2.MoveNext
    Loop
    rcTB2.Close
            
    On Error Resume Next
    Set rcTB2 = New ADODB.Recordset
    Err = 0
    Msg = "SELECT * FROM TB2"
    rcTB2.CursorLocation = adUseClient
    rcTB2.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
    If Err Then MsgBox "stop"
    rcTB2.MoveFirst
    
    Do While Not rcTB2.EOF
        strDocType = rcTB2.Fields("DocType")
        Select Case strDocType
            Case "01", "02"
                strMemoTB2 = rcTB2.Fields("MemoTB2")
                'search RFF+001: (start is position + 8)
                searchString = "RFF+001:"
                startIndex = InStr(strMemoTB2, searchString)
                If startIndex Then
                    strSub = Mid(strMemoTB2, startIndex + 8, 12)
                    
                    ' check for : in contract number
                    endIndex = InStr(strSub, ":")
                    If endIndex = 0 Then
                        ' check for ' in contract number
                        endIndex = InStr(strSub, "'")
                        If endIndex = 0 Then
                            If Len(strSub) = 12 Then
                                endIndex = 12
                            End If
                        End If
                    End If
                    strSub = Left(strSub, endIndex - 1)
                    bGet TABLE_CONTRACTS, 0, strSub
                    If Ktrl Then
                        rcTB2.Fields("Polis") = "999999999999"
                    Else
                        rcTB2.Fields("Polis") = strSub
                    End If
                Else
                    rcTB2.Fields("Polis") = "999999999999"
                End If
        End Select
        rcTB2.MoveNext
    Loop
    rcTB2.Close
    
    MsgBox "finished"
    
End Sub

Private Sub cmdEManueel_Click()

Dim MAPIString As String
Dim StringXGH As String
Dim StringXEH As String
Dim StringXET As String
Dim StringXGT As String
Dim MijNummer As String * 4

Dim Xarray() As String
Dim XGHarray() As String
Dim XEHarray() As String


MAPIString = RichTextBox1.Text
Xarray = Split(MAPIString, "'")
XGHarray = Split(Xarray(0) + "+", "+")
XEHarray = Split(Xarray(1) + "+", "+")

Dim XEHcode As String
Dim posSTR As Integer

    MijNummer = XGHarray(3)

XEHcode = Format(XEHarray(1), "00")
If InStr(MAPIString, "XEH+01") Then
    'XEH+01 = NIEUWE ZAAK OF WIJZIGING ?
    'MsgBox "RETOURBLOK worden enkel bewaard " & XGHarray(2), vbInformation
    'XEH+03 = KWIJTING!
    MsgBox "Uitwisseling Contante " & XGHarray(2), vbInformation
    EdiFactCONTANT MAPIString
    strA000RB = ""
    If InStr(MAPIString, "RFF+001:") = 0 Then
    Else
        Msg = Mid(MAPIString, InStr(MAPIString, "RFF+001:") + 8)
        Msg = Left(Msg, InStr(Msg, "'") - 1)
        If Msg <> XEHarray(4) Then
            On Error Resume Next
            Msg = Left(Msg, InStr(Msg, ":001") - 1)
        End If
        strA000RB = Msg
    End If
    GoSub VerwijderENslaOP
ElseIf InStr(MAPIString, "XEH+02") Then
    'Stop
    MsgBox "Opening schadedossier " & XGHarray(2), vbInformation
    'SchadeBeheer XEHarray(3), MAPIString, XGHarray(2)
    troelala$ = tb2Indent(MAPIString)
    Printer.Print troelala$
    Printer.EndDoc
    GoSub VerwijderENslaOP
ElseIf InStr(MAPIString, "XEH+03") Then
    'XEH+03 = KWIJTING!
    MsgBox "Uitwisseling Termijn " & XGHarray(2), vbInformation
    EdiFactTERMIJN MAPIString
    GoSub VerwijderENslaOP
ElseIf InStr(MAPIString, "XEH+07") Then
    'XEH+07 = REKENINGUITTREKSELS  ?
    MsgBox "Rekeninguittreksel " & XGHarray(2), vbInformation
    'Exit Sub
    EdiFactREKENINGUITTREKSEL MAPIString
    GoSub VerwijderENslaOP
Else
    MsgBox "Nog niets voorzien voor verzoek " & XEHarray(3) & " maatschappij " & XGHarray(2)
    
    GoSub VerwijderENslaOP
    Exit Sub
End If
Exit Sub

VerwijderENslaOP:
ProbeerNogEens:
KtrlBox = MsgBox("Bericht verwijderen.  Bent U zeker" & vbCrLf & "(Druk 'Ja' om te verwijderen, een copij wordt altijd bewaard!)", vbYesNoCancel + vbQuestion + vbDefaultButton3)
If KtrlBox = vbCancel Then
    GoTo ProbeerNogEens
ElseIf KtrlBox = vbNo Then
Else
    If XEHcode = "97" Then
        MsgBox "bericht met bijlagen nog verder te controleren. wordt nu verwijderd"
        On Local Error Resume Next
        'MPIBericht.Delete
        'VernieuwMAPIBerichten
        Exit Sub
    End If
    TLB_RECORD(TABLE_VARIOUS) = ""
    vBib TABLE_VARIOUS, vSet("L" + "CO" + Format(XGHarray(2), "0000"), 13), "v004"
    vBib TABLE_VARIOUS, Format(XGHarray(2), "0000"), "A010"
    vBib TABLE_VARIOUS, "E1", "v400"
    vBib TABLE_VARIOUS, vSet("30" + Format(XGHarray(2), "0000") + "E1", 20), "v005"
    vBib TABLE_VARIOUS, RichTextBox1.Text, "v132"
    vBib TABLE_VARIOUS, strA000RB, "A000"
      
    rsTB2.AddNew
    rsTB2("Mij").Value = Format(Left(XGHarray(2), 4), "0000")
    rsTB2("memoTB2").Value = RichTextBox1.Text
    rsTB2("DocType").Value = XEHcode
    rsTB2("Polis").Value = strA000RB
    'nu nog de datum op basis van doctype
    If XEHcode = "01" Then 'contract!
        posSTR = InStr(MAPIString, "DTM+008:")
        If posSTR = 0 Then
            'mogelijk nieuwe zaak dus ktrl op DTM+041:
            posSTR = InStr(MAPIString, "DTM+041:")
            If posSTR = 0 Then
                'verder zoeken!
                posSTR = InStr(MAPIString, "DTM+060:")
                If posSTR = 0 Then
                    'mogelijk beëindiging contract?
                    posSTR = InStr(MAPIString, "DTM+019:")
                    If posSTR = 0 Then
                        'nog verder zoeken!
                        MsgBox "Stop"
                    End If
                End If
            End If
        End If
        Msg = Mid(MAPIString, posSTR + 8)
        Msg = Left(Msg, 8)
        On Local Error Resume Next
        rsTB2("DatumAanvang").Value = CVDate(Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) & "/" & Mid(Msg, 5, 4))
    ElseIf XEHcode = "02" Then 'contract
        posSTR = InStr(MAPIString, "DTM+048:")
        If posSTR = 0 Then
            MsgBox "Stop"
        Else
            Msg = Mid(MAPIString, posSTR + 8)
            Msg = Left(Msg, 8)
        End If
        On Local Error Resume Next
        rsTB2("DatumAanvang").Value = CVDate(Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) & "/" & Mid(Msg, 5, 4))
    ElseIf XEHcode = "07" Then 'rekeninguittreksel?
        rsTB2("Polis").Value = "" 'zeker geen polisnummer dan!
        posSTR = InStr(MAPIString, "DTM+069:")
        If posSTR = 0 Then
            MsgBox "Stop"
        Else
            Msg = Mid(MAPIString, posSTR + 8)
            Msg = Left(Msg, 8)
        End If
        rsTB2("DatumAanvang").Value = CVDate(Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) & "/" & Mid(Msg, 5, 4))
    ElseIf XEHcode = "03" Then 'hernieuwingen ?
        rsTB2("Polis").Value = "" 'zeker geen polisnummer dan!
        posSTR = InStr(MAPIString, "DTM+005:")
        If posSTR = 0 Then
            MsgBox "Stop"
        Else
            Msg = Mid(MAPIString, posSTR + 8)
            Msg = Left(Msg, 6)
        End If
        On Local Error Resume Next
        rsTB2("DatumAanvang").Value = CVDate("01" & "/" & Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 4))
    Else
        MsgBox "Stop"
    End If
    rsTB2.Update
          
    bInsert TABLE_VARIOUS, 1
    If Ktrl Then
    Else
        On Local Error Resume Next
        'MPIBericht.Delete
        'VernieuwMAPIBerichten
    End If
End If
Return

End Sub

Private Sub cbMailManueel_Click()

Dim MAPIString As String
Dim StringXGH As String
Dim StringXEH As String
Dim StringXET As String
Dim StringXGT As String
Dim MijNummer As String * 4

Dim Xarray() As String
Dim XGHarray() As String
Dim XEHarray() As String


MAPIString = Me.rtbMail.Text
Xarray = Split(MAPIString, "'")
XGHarray = Split(Xarray(0) + "+", "+")
XEHarray = Split(Xarray(1) + "+", "+")

Dim XEHcode As String
Dim posSTR As Integer

    MijNummer = XGHarray(3)

XEHcode = Format(XEHarray(1), "00")
If InStr(MAPIString, "XEH+01") Then
    'XEH+01 = NIEUWE ZAAK OF WIJZIGING ?
    MsgBox "Uitwisseling Contante " & XGHarray(2), vbInformation
    EdiFactCONTANT MAPIString
    strA000RB = ""
    If InStr(MAPIString, "RFF+001:") = 0 Then
    Else
        Msg = Mid(MAPIString, InStr(MAPIString, "RFF+001:") + 8)
        Msg = Left(Msg, InStr(Msg, "'") - 1)
        If Msg <> XEHarray(4) Then
            On Error Resume Next
            Msg = Left(Msg, InStr(Msg, ":001") - 1)
        End If
        strA000RB = Msg
    End If
   
    'EdiFactTERMIJN MAPIString
    GoSub VerwijderENslaOP
ElseIf InStr(MAPIString, "XEH+02") Then
    'Stop
    MsgBox "Opening schadedossier " & XGHarray(2), vbInformation
    SchadeBeheer XEHarray(3), MAPIString, XGHarray(2)
    GoSub VerwijderENslaOP
ElseIf InStr(MAPIString, "XEH+03") Then
    'XEH+03 = KWIJTING!
    MsgBox "Uitwisseling Termijn " & XGHarray(2), vbInformation
    EdiFactTERMIJN MAPIString
    GoSub VerwijderENslaOP
ElseIf InStr(MAPIString, "XEH+07") Then
    'XEH+07 = REKENINGUITTREKSELS  ?
    MsgBox "Rekeninguittreksel " & XGHarray(2), vbInformation
    'Exit Sub
    EdiFactREKENINGUITTREKSEL MAPIString
    GoSub VerwijderENslaOP
Else
    MsgBox "Nog niets voorzien voor verzoek " & XEHarray(3) & " maatschappij " & XGHarray(2)
    Exit Sub
    GoSub VerwijderENslaOP
End If
Exit Sub

VerwijderENslaOP:
ProbeerNogEens:
KtrlBox = MsgBox("Bericht verwijderen.  Bent U zeker" & vbCrLf & "(Druk 'Ja' om te verwijderen, een copij wordt altijd bewaard!)", vbYesNoCancel + vbQuestion + vbDefaultButton3)
If KtrlBox = vbCancel Then
    GoTo ProbeerNogEens
ElseIf KtrlBox = vbNo Then
Else
    TLB_RECORD(TABLE_VARIOUS) = ""
    vBib TABLE_VARIOUS, vSet("L" + "CO" + Format(XGHarray(2), "0000"), 13), "v004"
    vBib TABLE_VARIOUS, Format(XGHarray(2), "0000"), "A010"
    vBib TABLE_VARIOUS, "E1", "v400"
    vBib TABLE_VARIOUS, vSet("30" + Format(XGHarray(2), "0000") + "E1", 20), "v005"
    vBib TABLE_VARIOUS, Me.rtbMail.Text, "v132"
    vBib TABLE_VARIOUS, strA000RB, "A000"
      
    rsTB2.AddNew
    rsTB2("Mij").Value = Format(Left(XGHarray(2), 4), "0000")
    rsTB2("memoTB2").Value = Me.rtbMail.Text
    rsTB2("DocType").Value = XEHcode
    rsTB2("Polis").Value = strA000RB
    'nu nog de datum op basis van doctype
    If XEHcode = "01" Then 'contract!
        posSTR = InStr(MAPIString, "DTM+008:")
        If posSTR = 0 Then
            'mogelijk nieuwe zaak dus ktrl op DTM+041:
            posSTR = InStr(MAPIString, "DTM+041:")
            If posSTR = 0 Then
                'verder zoeken!
                posSTR = InStr(MAPIString, "DTM+060:")
                If posSTR = 0 Then
                    'mogelijk beëindiging contract?
                    posSTR = InStr(MAPIString, "DTM+019:")
                    If posSTR = 0 Then
                        'nog verder zoeken!
                        MsgBox "Stop"
                    End If
                End If
            End If
        End If
        Msg = Mid(MAPIString, posSTR + 8)
        Msg = Left(Msg, 8)
        rsTB2("DatumAanvang").Value = CVDate(Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) & "/" & Mid(Msg, 5, 4))
    ElseIf XEHcode = "02" Then 'contract
        posSTR = InStr(MAPIString, "DTM+048:")
        If posSTR = 0 Then
            MsgBox "Stop"
        Else
            Msg = Mid(MAPIString, posSTR + 8)
            Msg = Left(Msg, 8)
        End If
        rsTB2("DatumAanvang").Value = CVDate(Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) & "/" & Mid(Msg, 5, 4))
    ElseIf XEHcode = "07" Then 'rekeninguittreksel?
        rsTB2("Polis").Value = "" 'zeker geen polisnummer dan!
        posSTR = InStr(MAPIString, "DTM+069:")
        If posSTR = 0 Then
            MsgBox "Stop"
        Else
            Msg = Mid(MAPIString, posSTR + 8)
            Msg = Left(Msg, 8)
        End If
        rsTB2("DatumAanvang").Value = CVDate(Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) & "/" & Mid(Msg, 5, 4))
    ElseIf XEHcode = "03" Then 'hernieuwingen ?
        rsTB2("Polis").Value = "" 'zeker geen polisnummer dan!
        posSTR = InStr(MAPIString, "DTM+005:")
        If posSTR = 0 Then
            MsgBox "Stop"
        Else
            Msg = Mid(MAPIString, posSTR + 8)
            Msg = Left(Msg, 6)
        End If
        rsTB2("DatumAanvang").Value = CVDate("01" & "/" & Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 4))
    Else
        MsgBox "Stop"
    End If
    rsTB2.Update
          
    bInsert TABLE_VARIOUS, 1
    If Ktrl Then
    Else
        On Local Error Resume Next
        'Me.MPIBericht.Delete
        'Me.refreshMPIBericht
    End If
End If
Return

End Sub

Private Sub cmdPManueel_Click()

Dim MAPIString As String
Dim StringXGH As String
Dim StringXEH As String
Dim StringXET As String
Dim StringXGT As String
Dim MijNummer As String * 4

Dim Xarray() As String
Dim XGHarray() As String
Dim XEHarray() As String


MAPIString = RichTextBox1.Text
Xarray = Split(MAPIString, "'")
XGHarray = Split(Xarray(0) + "+", "+")
XEHarray = Split(Xarray(1) + "+", "+")
    MijNummer = XGHarray(3)

If InStr(MAPIString, "XGH") Then
    StringXGH = Mid(MAPIString, 1, InStr(MAPIString, "'") - 1)
    MAPIString = Mid(MAPIString, InStr(MAPIString, "'") + 1)
End If
If InStr(MAPIString, "XEH") Then
    StringXEH = Mid(MAPIString, 1, InStr(MAPIString, "'") - 1)
    MAPIString = Mid(MAPIString, InStr(MAPIString, "'") + 1)
    'XEH+02+1+205, '205' is vermoedelijk opening schadedossier ?
    'XEH+xx+x+0103, '0103' is vermoedelijk antwoordrecord nieuwe polis recht ?
    'XEH+xx+x+304, '304' is vermoedelijk termijn-80 ?
End If
If InStr(MAPIString, Chr(254)) Then
    'AS/1 termijn-80 bericht of schadebericht
    'userarea
    Select Case Format(XEHarray(3), "0000")
        Case "0101"
            'Consultatie BRAND ?
            MsgBox "Stop raadpleging BRAND " & " maatschappij " & XGHarray(2), vbInformation
            GoSub VerwijderENslaOP
            
        Case "0103"
            'Bloc retour BRAND ?
            MsgBox "Stop bloc retour BRAND? " & XEHarray(3) & " maatschappij " & XGHarray(2), vbQuestion
            GoSub VerwijderENslaOP
            
        Case "0108"
            'plaatsingsmandaat
            Msg = "Polisnummer " & Trim$(Mid(MAPIString, 1, 12)) & vbCr
            Msg = Msg & "Agentnummer " & Trim$(Mid(MAPIString, 13, 8)) & vbCr
            Msg = Msg & "Maatschappij " & XGHarray(2) & vbCr
            Msg = Msg & "Klant " & Trim$(Mid(MAPIString, 21, 15)) & vbCr
            Msg = Msg & "Einddatum " & Trim$(Mid(MAPIString, 66, 8))
            MsgBox Msg, vbExclamation, "Ontnomen mandaat !"
            GoSub VerwijderENslaOP
            
        Case "0205"
            'Schades, 205 = bericht opening dossier
            'Schades, 206 = regelingsbericht
            Debug.Assert False
            aa = Mid(MAPIString, 1, InStr(MAPIString, Chr(254)) - 1)
            MAPIString = Mid(MAPIString, InStr(MAPIString, Chr(254)) + 2)
    
            If InStr(MAPIString, "XET") Then
                StringXET = Mid(MAPIString, 1, InStr(MAPIString, "'") - 1)
                MAPIString = Mid(MAPIString, InStr(MAPIString, "'") + 1)
            End If
            If InStr(MAPIString, "XGT") Then
                StringXGT = Mid(MAPIString, 1, InStr(MAPIString, "'") - 1)
                MAPIString = Mid(MAPIString, InStr(MAPIString, "'") + 1)
            End If
            If MAPIString = "" Then
                'schades (voor 99,99 %)
                'SchadeBeheer2 Format(XGHarray(2), "0000")
                GoSub VerwijderENslaOP
            Else
                MsgBox "schade nog iets anders: stop"
            End If
            
        Case "0206" 'schade regelingsbericht ?
            MsgBox "Schade regelingsbericht", vbQuestion
            Debug.Assert False
            aa = Mid(MAPIString, 1, InStr(MAPIString, Chr(254)) - 1)
            MAPIString = Mid(MAPIString, InStr(MAPIString, Chr(254)) + 2)
    
            If InStr(MAPIString, "XET") Then
                StringXET = Mid(MAPIString, 1, InStr(MAPIString, "'") - 1)
                MAPIString = Mid(MAPIString, InStr(MAPIString, "'") + 1)
            End If
            If InStr(MAPIString, "XGT") Then
                StringXGT = Mid(MAPIString, 1, InStr(MAPIString, "'") - 1)
                MAPIString = Mid(MAPIString, InStr(MAPIString, "'") + 1)
            End If
            If MAPIString = "" Then
                'schades (voor 99,99 %)
                'SchadeBeheer2 Format(XGHarray(2), "0000")
                GoSub VerwijderENslaOP
            Else
                MsgBox "schade nog iets anders: stop"
            End If
            
                        
        Case "0304"
            'Uitwisseling 80
            aa = Mid(MAPIString, 1, InStr(MAPIString, Chr(254)) - 1)
            MAPIString = Mid(MAPIString, InStr(MAPIString, Chr(254)) + 2)
    
            If InStr(MAPIString, "XET") Then
                StringXET = Mid(MAPIString, 1, InStr(MAPIString, "'") - 1)
                MAPIString = Mid(MAPIString, InStr(MAPIString, "'") + 1)
            End If
            If InStr(MAPIString, "XGT") Then
                StringXGT = Mid(MAPIString, 1, InStr(MAPIString, "'") - 1)
                MAPIString = Mid(MAPIString, InStr(MAPIString, "'") + 1)
            End If
            If MAPIString = "" Then
                'AS/1 termijn-80 bericht (voor 99,99 %)
                MsgBox "Uitwisseling 80 voor " & XGHarray(2), vbInformation
                Borderel Format(XGHarray(2), "0000"), aa
                GoSub VerwijderENslaOP
            Else
                MsgBox "nog iets anders: stop"
            End If
        
        Case Else
            MsgBox "Nog niets voorzien voor verzoek " & XEHarray(3) & " maatschappij " & XGHarray(2)
            GoSub VerwijderENslaOP
    End Select
Else
    MsgBox "Op dit ogenblik nog geen verwerkingsmogelijkheid beschikbaar voor dit bericht", vbInformation
End If
Exit Sub

VerwijderENslaOP:
ProbeerNogEens:
KtrlBox = MsgBox("Bericht verwijderen.  Bent U zeker" & vbCrLf & "(Druk 'Ja' om te verwijderen, een copij wordt altijd bewaard!)", vbYesNoCancel + vbQuestion + vbDefaultButton3)
If KtrlBox = vbCancel Then
    GoTo ProbeerNogEens
ElseIf KtrlBox = vbNo Then
Else
    TLB_RECORD(TABLE_VARIOUS) = ""
    vBib TABLE_VARIOUS, vSet("L" + "CO" + Format(XGHarray(2), "0000"), 13), "v004"
    vBib TABLE_VARIOUS, Format(XGHarray(2), "0000"), "A010"
    vBib TABLE_VARIOUS, "P1", "v400"
    vBib TABLE_VARIOUS, vSet("30" + Format(XGHarray(2), "0000") + "P1", 20), "v005"
    vBib TABLE_VARIOUS, RichTextBox1.Text, "v132"
    bInsert TABLE_VARIOUS, 1
    If Ktrl Then
    Else
        'MPIBericht.Delete
        'VernieuwMAPIBerichten
    End If
End If
Return

End Sub

Private Sub Form_Load()

    Dim COUNT_TO As Integer
    
    Screen.MousePointer = vbHourglass
    Set rsTB2 = New ADODB.Recordset
    rsTB2.CursorLocation = adUseClient
    If rsTB2.State = adStateOpen Then
        rsTB2.Close
    End If
    Msg = "SELECT TOP 1 * FROM TB2"
    rsTB2.Open Msg, adntDB, adOpenDynamic, adLockOptimistic, adCmdText
           
    gboDIR = Trim$(LaadTekst(App.Title, "gboDIR"))
    gboFILE = "PRENOT.GBO"
    If gboDIR = "" Then
        Me.cbGBO.Enabled = False
    Else
        Me.cbGBO.Enabled = True
        If ScrLeesBestandAlleTekst(gboCPstr, gboDIR & gboFILE) Then
            Me.RichTextBox2.Text = gboCPstr
        Else
            Me.RichTextBox2.Text = ""
        End If
    End If
    Label1(4).Caption = gboDIR & gboFILE
    
    aswebDIR = Trim$(LaadTekst(App.Title, "aswebDIR"))
    aswebFILE = "AS1.EDI"
    If aswebDIR = "" Then
        Me.cbASWEB.Enabled = False
    Else
        Me.cbASWEB.Enabled = True
        If ScrLeesBestandAlleTekst(aswebCPstr, aswebDIR & aswebFILE) Then
            Me.RichTextBox3.Text = aswebCPstr
        Else
            Me.RichTextBox3.Text = ""
        End If
    End If
    Label1(5).Caption = aswebDIR & aswebFILE
    
GridTextPolis = ""
cbAS2BerichtType.Clear
cbAS2BerichtType.AddItem "IPC.AS2.E1: IPC Edifact Bericht (AS2)"
cbAS2BerichtType.AddItem "IPC.AS2.P1: IPC positioneel Bericht (AS2)"
cbAS2BerichtType.AddItem "IPM.NOTE.ASSURNET: IPM Vrij Bericht Assurnet"

cbAS2BerichtTypeHistoriek.Clear
cbAS2BerichtTypeHistoriek.AddItem "E1: IPC Edifact Bericht (AS2)"
cbAS2BerichtTypeHistoriek.AddItem "P1: IPC positioneel Bericht (AS2)"
cbAS2BerichtTypeHistoriek.AddItem "IPM.NOTE.ASSURNET: IPM Vrij Bericht Assurnet"

InstallMij

On Error Resume Next
MPISessie.DownLoadMail = True
MPISessie.SignOn
If Err Then
    SnelHelpPrint "GBO Services beschikbaar - AS2Dispatch niet beschikbaar", False
    'MsgBox "GBO Services beschikbaar" & vbCrLf & "AS2Dispatch niet beschikbaar" & vbCrLf & Error, vbInformation
    cbAS2BerichtType.Enabled = False
Else
    MPIBericht.SessionID = MPISessie.SessionID
    cbAS2BerichtType.ListIndex = 0
End If

On Error Resume Next
Me.SetFocus
Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
Me.MPIBericht.Save
MPISessie.SignOff

End Sub


Private Sub PrintVrijBericht()
Dim MijNaam As String
Dim PolisOfSchadenummer As String * 12
Dim TypeBericht As String
Dim Beheerder As String

bGet TABLE_SUPPLIERS, 0, vSet("CO" + Mid(A, 58, 4), 12)
If Ktrl Then
    MijNaam = "[" + Mid(A, 58, 4) + " Nog in te brengen als leverancier!!!" + "]"
    FVT(TABLE_SUPPLIERS, 0) = Space$(6)
    FVT(TABLE_SUPPLIERS, 1) = Space$(35)
Else
    RecordToVeld TABLE_SUPPLIERS
    MijNaam = "[" + RTrim$(FVT(TABLE_SUPPLIERS, 1)) + "]"
End If

PolisOfSchadenummer = LTrim$(RTrim$(Mid(A, 106, 12)))
TLB_RECORD(TABLE_CUSTOMERS) = ""

KlantNaam1 = "Polis niet (meer) aanwezig !"
Klantnaam2 = ""
KlantStraat = ""
KlantPlaats = ""
KlantTelefoon = ""
KlantFax = ""
Select Case Mid(A, 105, 1)
    Case "P"
        TypeBericht = "Produktie "
        bGet TABLE_CONTRACTS, 0, PolisOfSchadenummer
        If Ktrl Then
            DetailPolis1 = ""
            DetailPolis2 = ""
        Else
            RecordToVeld TABLE_CONTRACTS
            DetailPolis1 = vBibTekst(TABLE_CONTRACTS, "#vs99 #")
            DetailPolis2 = vBibTekst(TABLE_CONTRACTS, "#vs98 #")
            bGet TABLE_CUSTOMERS, 0, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12)
            If Ktrl Then
                MsgBox "Stop tijdens zoeken klant"
            Else
                RecordToVeld TABLE_CUSTOMERS
                KlantNaam1 = vBibTekst(TABLE_CUSTOMERS, "#A100 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A101 #")
                Klantnaam2 = vBibTekst(TABLE_CUSTOMERS, "#A125 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A127 #")
                KlantStraat = vBibTekst(TABLE_CUSTOMERS, "#A104 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A105 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A106 #")
                KlantPlaats = vBibTekst(TABLE_CUSTOMERS, "#A109 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A107 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A108 #")
                KlantTelefoon = vBibTekst(TABLE_CUSTOMERS, "#A10A #")
                KlantFax = vBibTekst(TABLE_CUSTOMERS, "#vs02 #")
            End If
        End If
    Case "S"
        TypeBericht = "Schade "
    Case Else
        TypeBericht = "Andere "
        Select Case Mid(A, 23, 1)
            Case "2"
                TypeBericht = TypeBericht + ", vanwege Maatschappij"
            Case "4"
                TypeBericht = TypeBericht + ", vanwege ASSURNET"
            Case "6"
                TypeBericht = TypeBericht + ", vanwege INFORMEX"
        End Select
End Select
Beheerder = Mid(A, 121, 10)

    PAGE_COUNTER = 0
    Printer.FontName = "Courier New"
    Printer.FontSize = 12
    Printer.FontBold = True
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print UCase(Mid(Mim.Caption, InStr(Mim.Caption, "[")));
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print "marINTEGRAAL NT : ";
    Printer.FontBold = True
        Printer.Print "Vrij bericht : "; TypeBericht;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf;
        Printer.Print Tab(1); "Beheerder     : ";
    Printer.FontBold = True
        Printer.Print Beheerder;
    Printer.FontBold = False
        Printer.Print Tab(38); "Dossiernummer : ";
    Printer.FontBold = True
        Printer.Print PolisOfSchadenummer;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Tab(1); "Maatschappij  : ";
    Printer.FontBold = True
        Printer.Print Mid(FVT(TABLE_SUPPLIERS, 0), 3, 4); " "; FVT(TABLE_SUPPLIERS, 1);
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf;

    If Left(TypeBericht, 6) = "Andere" Then
        GoTo JumpAndere
    End If
        
        Printer.Print Tab(1); "Klanten       : ";
    Printer.FontBold = True
        Printer.Print KlantNaam1; " "; vbCrLf;
        Printer.Print Tab(1); "                "; Klantnaam2;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Tab(1); "Adres         : ";
    Printer.FontBold = True
        Printer.Print KlantStraat; vbCrLf;
        Printer.Print Tab(1); "                "; KlantPlaats;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Tab(1); "Telefoon      : ";
    Printer.FontBold = True
        Printer.Print KlantTelefoon;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Tab(1); "Fax           : ";
    Printer.FontBold = True
        Printer.Print KlantFax;
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf;
    Printer.FontBold = True
        Printer.Print DetailPolis1;
    Printer.FontBold = False
        Printer.Print vbCrLf;
    Printer.FontBold = True
        Printer.Print DetailPolis2;
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf;

JumpAndere:
        Printer.Print Tab(1); "Onderwerp : ";
    Printer.FontBold = True
        Printer.Print Mid(A, 65, 40);
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf; vbCrLf;
    
    Do While aa <> ""
        Printer.Print Left(aa, 79)
        aa = Right(aa, Len(aa) - 79)
    Loop

        Printer.Print vbCrLf; vbCrLf; "Datum en uur vandaag             : ";
    Printer.FontBold = True
        Printer.Print Format(Now, "dd/mm/yy") + " " + Format(Now, "hh:mm:ss");
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print "Datum en uur opstelling bericht  : ";
    Printer.FontBold = True
        Printer.Print Mid(A, 36, 2) + "/" + Mid(A, 34, 2) + "/" + Mid(A, 32, 2) + " " + Mid(A, 38, 2) + ":" + Mid(A, 40, 2) + ":" + Mid(A, 42, 2);
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print "Datum en uur verzending ASSURNET : ";
    Printer.FontBold = True
        Printer.Print Mid(A, 48, 2) + "/" + Mid(A, 46, 2) + "/" + Mid(A, 44, 2) + " " + Mid(A, 50, 2) + ":" + Mid(A, 52, 2) + ":" + Mid(A, 54, 2);
    Printer.FontBold = False
        Printer.Print vbCrLf;
    Printer.NewPage
    Printer.EndDoc

End Sub

Private Sub RefreshControls()

If Mid(A, 17, 3) = "913" Then
    Exit Sub
End If

UserInfo(0).Caption = Mid(A, 17, 3)
bGet TABLE_SUPPLIERS, 0, vSet("CO" + Mid(A, 58, 4), 12)
If Ktrl Then
    FVT(TABLE_SUPPLIERS, 1) = "Maatschappij nog in te voegen"
Else
    RecordToVeld TABLE_SUPPLIERS
End If
UserInfo(1).Caption = FVT(TABLE_SUPPLIERS, 0) + " : " + FVT(TABLE_SUPPLIERS, 1)

End Sub

Private Sub SchadeBeheer(strType As String, strBericht As String, strMij As String)

Dim PolisNummer As String * 12
Dim Dossier As String * 12
Dim MijNaam As String
Dim TypeBericht As String
Dim TeleBibTekst As String
Dim Msg As String
Dim OptieTxt As String
Dim fKtrl As Integer
Dim T As Integer
Dim DataPos As Integer
Dim DataLen As Integer
Dim Beheerder As String


Dim Xarray() As String
Dim XEHarray() As String
Dim XGHarray() As String
Dim MOAarray() As String
Dim DTMarray() As String

Dim typePTY As String
Dim Telhier As Integer

typePTY = ""
Xarray = Split(strBericht, "'")
XGHarray = Split(Xarray(0), "+")

lijnTeller = 0

bGet TABLE_SUPPLIERS, 0, vSet("CO" & strMij, 12)
If Ktrl Then
    MijNaam = "[" & strMij & " Niet in database als leverancier!!!" & "]"
    FVT(TABLE_SUPPLIERS, 0) = Space$(6)
    FVT(TABLE_SUPPLIERS, 1) = Space$(30)
Else
    RecordToVeld TABLE_SUPPLIERS
    MijNaam = "[" & Trim$(FVT(TABLE_SUPPLIERS, 1)) & "]"
End If

Telhier = 0
Do While Telhier <= UBound(Xarray) - 1
        'Dossier = LTrim$(RTrim$(Mid(A, 231, 12)))
        'polisNummer = LTrim$(RTrim$(Mid(A, 142, 12)))
        'Beheerder = Mid(A, 121, 10)

    Select Case Left(Xarray(Telhier), 3)
        Case "DTM"
            'Stop
            DTMarray = Split(Xarray(Telhier), ":")
            If Left(Xarray(Telhier), 8) = "DTM+005:" Then
                If DTMarray(UBound(DTMarray)) = "001" Then
                    MaandVerwerking = Mid(Xarray(Telhier), 11, 2) 'boekMAAND
                    BoekjaarKontrole = Mid(Xarray(Telhier), 13, 4) 'boekJAAR
                ElseIf DTMarray(UBound(DTMarray)) = "005" Then
                    MaandVerwerking = Mid(Xarray(Telhier), 9, 2)    'boekMAAND
                    BoekjaarKontrole = Mid(Xarray(Telhier), 11, 4)  'boekJAAR
                Else
                    MsgBox "stoppen en controleren a.u.b"
                End If
            ElseIf Left(Xarray(Telhier), 8) = "DTM+004:" Then
                DagKwijting = Mid(DTMarray(1), 1, 2)
                MaandKwijting = Mid(DTMarray(1), 3, 2)
                DatumKwijting = DagKwijting & "/" & MaandKwijting & "/" & BoekjaarKontrole 'Mid(DTMarray(1), 5, 4)
            End If
        
        Case "RFF"
            'Stop
            Select Case Mid(Xarray(Telhier), 5, 3)
                Case "001"  'polisnummer (eerst klant, dan tpartij)
                   ' Stop
                    If PolisNummer = "" Then
                        PolisNummer = Mid(Xarray(Telhier), 9)
                    Else
                        'Stop
                    End If
           
                Case "029"  'schadenummer maar opgelet, daarna ook ref.agent!
                    'Stop
                Case "010"  'plaatnummer
                    'Stop
                Case "011"  'chassis
            Case Else
                'MsgBox Xarray(Telhier)
            End Select
            
           
        Case "MOA"
            'Stop
            MOAarray = Split(Xarray(Telhier), ":")
            
            Select Case MOAarray(0)
                Case "MOA+012"  'Totale premie
                    If MOAarray(2) = "EUR" Then
                        If MOAarray(3) = 2 Then
                            Bedrag0 = Val(MOAarray(1)) / 100
                        ElseIf MOAarray(3) = 0 Then
                            MsgBox "stop"
                            Bedrag0 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    ElseIf MOAarray(2) = "BEF" Then
                        Bedrag0 = Val(MOAarray(1))
                    Else
                        MsgBox "Stop"
                    End If
                
                Case "MOA+013"  'netto premie BA auto
                    If MOAarray(2) = "EUR" Then
                        If MOAarray(3) = 2 Then
                            Bedrag1 = Val(MOAarray(1)) / 100
                        ElseIf MOAarray(3) = 0 Then
                            MsgBox "stop"
                            Bedrag1 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    ElseIf MOAarray(2) = "BEF" Then
                        Bedrag1 = Val(MOAarray(1))
                    Else
                        MsgBox "Stop"
                    End If
                
                Case "MOA+015"  'commissie
                    If MOAarray(2) = "EUR" Then
                        If MOAarray(3) = 2 Then
                            Bedrag9 = Val(MOAarray(1)) / 100
                        ElseIf MOAarray(3) = 0 Then
                            MsgBox "stop"
                            Bedrag9 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    ElseIf MOAarray(2) = "BEF" Then
                        Bedrag9 = Val(MOAarray(1))
                    Else
                        MsgBox "Stop"
                    End If
                                   
                   
                Case "MOA+016"  'Lastenbedrag BA auto
                    If MOAarray(2) = "EUR" Then
                        If MOAarray(3) = 2 Then
                            Bedrag2 = Val(MOAarray(1)) / 100
                        ElseIf MOAarray(3) = 0 Then
                            MsgBox "stop"
                            Bedrag2 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    ElseIf MOAarray(2) = "BEF" Then
                        Bedrag2 = Val(MOAarray(1))
                    Else
                        MsgBox "Stop"
                    End If
                                    
                Case "MOA+017"  'Splitsingskosten? + enkel BA auto ?
                    'MsgBox "nog manueel te controleren, voor AG reeds in orde", vbExclamation
                    If MOAarray(2) = "EUR" Then
                        If MOAarray(3) = 2 Then
                            Bedrag2 = Bedrag2 + Val(MOAarray(1)) / 100
                        ElseIf MOAarray(3) = 0 Then
                            MsgBox "stop"
                            Bedrag2 = Bedrag2 + Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    ElseIf MOAarray(2) = "BEF" Then
                        Bedrag2 = Bedrag2 + Val(MOAarray(1))
                    Else
                        MsgBox "Stop"
                    End If
                                    
                'ANDERE DAN AUTO !
                Case "MOA+097"  'premie andere waarborgen dan BA auto
                    If MOAarray(2) = "EUR" Then
                        If MOAarray(3) = 2 Then
                            Bedrag4 = Val(MOAarray(1)) / 100
                        ElseIf MOAarray(3) = 0 Then
                            MsgBox "stop"
                            Bedrag4 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    ElseIf MOAarray(2) = "BEF" Then
                        Bedrag4 = Val(MOAarray(1))
                    Else
                        MsgBox "Stop"
                    End If
                
                Case "MOA+098"  'taksen en kosten andere dan BA auto
                    If MOAarray(2) = "EUR" Then
                        If MOAarray(3) = 2 Then
                            Bedrag5 = Val(MOAarray(1)) / 100
                        ElseIf MOAarray(3) = 0 Then
                            MsgBox "stop"
                            Bedrag5 = Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    ElseIf MOAarray(2) = "BEF" Then
                        Bedrag5 = Val(MOAarray(1))
                    Else
                        MsgBox "Stop"
                    End If

                Case "MOA+100"  'Kwitantiekosten
                    MsgBox "nog manueel te controleren, voor MAURETUS reeds in orde", vbExclamation
                    If MOAarray(2) = "EUR" Then
                        If MOAarray(3) = 2 Then
                            Bedrag2 = Bedrag2 + Val(MOAarray(1)) / 100
                        ElseIf MOAarray(3) = 0 Then
                            MsgBox "stop"
                            Bedrag2 = Bedrag2 + Val(MOAarray(1))
                        Else
                            MsgBox "Stop"
                        End If
                    ElseIf MOAarray(2) = "BEF" Then
                        Bedrag2 = Bedrag2 + Val(MOAarray(1))
                    Else
                        MsgBox "Stop"
                    End If
                    
                Case "MOA+105"  'Andere aan netto premie toe te voegen kosten
                
                    
                
                Case Else
                    'MsgBox "MOA-stop voor " & MOAarray(0)
            End Select
            
        Case "NME"
            'Stop
            'TempoNaamKlant = Mid(Xarray(Telhier), 9)
        
        Case "PTY"
            'Stop
            
                
        Case "ATT"
            'Stop
            If Xarray(Telhier) = "ATT+A600+3" Then 'Inning Maatschappij
                MsgBox "Kwitantie voor polis " & Trim$(PolisNummer) & " staat inning Maatschappij", vbInformation
                
                Do While Xarray(Telhier) <> "XET+03"
                    Telhier = Telhier + 1
                Loop
            ElseIf Left(Xarray(Telhier), 8) = "ATT+5300" Then 'bonus malus
                HuidigeBM = Mid(Xarray(Telhier), 10, 2)
                'MsgBox "bonus malus controle: " & HuidigeBM & " " & TempoNaamKlant
            End If
            
        Case "XET" 'Afsluiten/lijn bewaren
            'stop en afsluiten
            
        Case Else
        
            'MsgBox Xarray(Telhier)

            'voorlopig over te slaan...
      
    End Select
    Telhier = Telhier + 1
Loop

bGet TABLE_VARIOUS, 1, "24" + Mid(A, 58, 4) + Dossier
If Ktrl Then
    fKtrl = 99
    TLB_RECORD(TABLE_VARIOUS) = ""
Else
    fKtrl = 100
    RecordToVeld TABLE_VARIOUS
End If

TeleBibSchade
TypeBericht = "SCHADE BERICHT/REGELINGSBERICHT"

'Stop


bGet TABLE_CONTRACTS, 0, PolisNummer
If Ktrl Then
    DetailPolis1 = ""
    DetailPolis2 = ""
Else
    RecordToVeld TABLE_CONTRACTS
    DetailPolis1 = vBibTekst(TABLE_CONTRACTS, "#vs99 #")
    DetailPolis2 = vBibTekst(TABLE_CONTRACTS, "#vs98 #")
End If
Do While Ktrl <> 0
          Msg = "Polis " + PolisNummer + " bestaat niet eens." + vbCrLf
    Msg = Msg + "Breng alles straks manueel in a.u.b." + vbCrLf
    Msg = Msg + Mid(A, 186, 35)
    MsgBox Msg, 0, "Gebruikersfout"
    GoTo AfTeDrukken
Loop
bGet TABLE_CUSTOMERS, 0, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12)
If Ktrl Then
    MsgBox "Klant niet te vinden voor polis :" + PolisNummer
Else
    RecordToVeld TABLE_CUSTOMERS
    KlantNaam1 = vBibTekst(TABLE_CUSTOMERS, "#A100 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A101 #")
    Klantnaam2 = vBibTekst(TABLE_CUSTOMERS, "#A125 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A127 #")
    KlantStraat = vBibTekst(TABLE_CUSTOMERS, "#A104 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A105 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A106 #")
    KlantPlaats = vBibTekst(TABLE_CUSTOMERS, "#A109 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A107 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A108 #")
    KlantTelefoon = vBibTekst(TABLE_CUSTOMERS, "#A10A #")
    KlantFax = vBibTekst(TABLE_CUSTOMERS, "#vs02 #")
End If

T = 0
Do While TLBCode(T) <> Space$(10)
    On Local Error Resume Next
    vBib TABLE_VARIOUS, Mid(A, TLBPos(T), TLBLengte(T)), Mid(TLBCode(T), 5, 5)
    T = T + 1
Loop
vBib TABLE_VARIOUS, vSet("K" + vBibTekst(TABLE_CONTRACTS, "#A110 #"), 13), "v004"
vBib TABLE_VARIOUS, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12), "A110"
vBib TABLE_VARIOUS, Mid(A, 58, 4), "A010"
vBib TABLE_VARIOUS, PolisNummer, "A000"
vBib TABLE_VARIOUS, Dossier, "C000"
vBib TABLE_VARIOUS, vSet("24" + Mid(A, 58, 4) + Dossier, 20), "v005"
If fKtrl = 99 Then
    bInsert TABLE_VARIOUS, 1
Else
    bUpdate TABLE_VARIOUS, 1
End If
AfTeDrukken:
    PAGE_COUNTER = 0
    Printer.FontName = "Courier New"
    Printer.FontSize = 12
    Printer.FontBold = True
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print UCase(Mid(Mim.Caption, InStr(Mim.Caption, "[")));
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print "marINTEGRAAL NT : ";
    Printer.FontBold = True
        Printer.Print TypeBericht;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf;
        Printer.Print Tab(1); "Beheerder     : ";
    Printer.FontBold = True
        Printer.Print Beheerder;
    Printer.FontBold = False
        Printer.Print Tab(38); "Dossiernummer : ";
    Printer.FontBold = True
        Printer.Print Dossier;
    Printer.FontBold = False
        Printer.Print Tab(38); "PolisNummer : ";
    Printer.FontBold = True
        Printer.Print PolisNummer;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Tab(1); "Maatschappij  : ";
    Printer.FontBold = True
        Printer.Print Mid(FVT(TABLE_SUPPLIERS, 0), 3, 4); " "; FVT(TABLE_SUPPLIERS, 1);
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf;
        Printer.Print Tab(1); "Klanten       : ";
    Printer.FontBold = True
        Printer.Print KlantNaam1; " "; vbCrLf;
        Printer.Print Tab(1); "                "; Klantnaam2;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Tab(1); "Adres         : ";
    Printer.FontBold = True
        Printer.Print KlantStraat; vbCrLf;
        Printer.Print Tab(1); "                "; KlantPlaats;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Tab(1); "Telefoon      : ";
    Printer.FontBold = True
        Printer.Print KlantTelefoon;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Tab(1); "Fax           : ";
    Printer.FontBold = True
        Printer.Print KlantFax;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf;
    Printer.FontBold = True
        Printer.Print DetailPolis1;
    Printer.FontBold = False
        Printer.Print vbCrLf;
    Printer.FontBold = True
        Printer.Print DetailPolis2;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf;

    T = 0
    Do While TLBCode(T) <> Space$(10)
        If RTrim$(vBibTekst(TABLE_VARIOUS, "#" + Mid(TLBCode(T), 5, 5) + "#")) = "" Then
        Else
            OptieTxt = vBibTekst(TABLE_VARIOUS, "#" + Mid(TLBCode(T), 5, 5) + "#")
            Printer.FontBold = False
            Printer.Print Tab(1); TLBTekst(T);
            Printer.FontBold = True
            If Mid(TLBCode(T), 2, 2) <> "  " Then
                Printer.Print Tab(40); fmarBoxText(Mid(TLBCode(T), 2, 2), "2", OptieTxt);
            Else
                If Len(OptieTxt) > 35 Then
                    Print vbCrLf;
                    Printer.Print Tab(40); OptieTxt;
                Else
                    Printer.Print Tab(40); OptieTxt;
                End If
            End If
            Printer.Print vbCrLf;
        End If
        T = T + 1
    Loop
    
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf; "Datum en uur vandaag             : ";
    Printer.FontBold = True
        Printer.Print Format(Now, "dd/mm/yy") + " " + Format(Now, "hh:mm:ss");
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print "Datum en uur opstelling bericht  : ";
    Printer.FontBold = True
        Printer.Print Mid(A, 36, 2) + "/" + Mid(A, 34, 2) + "/" + Mid(A, 32, 2) + " " + Mid(A, 38, 2) + ":" + Mid(A, 40, 2) + ":" + Mid(A, 42, 2);
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print "Datum en uur verzending ASSURNET : ";
    Printer.FontBold = True
        Printer.Print Mid(A, 48, 2) + "/" + Mid(A, 46, 2) + "/" + Mid(A, 44, 2) + " " + Mid(A, 50, 2) + ":" + Mid(A, 52, 2) + ":" + Mid(A, 54, 2);
    Printer.FontBold = False
        Printer.Print vbCrLf;
    Printer.NewPage
    Printer.EndDoc

End Sub

Private Sub TeleBibSchade()

TLBCode(0) = "N01 " + "A500"
TLBTekst(0) = "Code Tak"
TLBType(0) = "N"

TLBCode(1) = "A   " + "A000"
TLBTekst(1) = "Polis 1"
TLBType(1) = "A"

TLBCode(2) = "A   " + "C041"
TLBTekst(2) = "Polis 2"
TLBType(2) = "A"

TLBCode(3) = "A   " + "C043"
TLBTekst(3) = "Polis 3"
TLBType(3) = "A"

TLBCode(4) = "A   " + "A400"
TLBTekst(4) = "Producent Nummer"
TLBType(4) = "A"

TLBCode(5) = "A   " + "A100"
TLBTekst(5) = "Verzekeringsnemer Naam"
TLBType(5) = "A"

TLBCode(6) = "A   " + "A110"
TLBTekst(6) = "Referte Producent (polis)"
TLBType(6) = "A"

TLBCode(7) = "A   " + "C000"
TLBTekst(7) = "Schadenummer maatschappij"
TLBType(7) = "A"

TLBCode(8) = "A   " + "C010"
TLBTekst(8) = "Referte Producent (schadegeval)"
TLBType(8) = "A"

TLBCode(9) = "N   " + "C200"
TLBTekst(9) = "Schadegeval datum"
TLBType(9) = "N"

TLBCode(10) = "A   " + "C002"
TLBTekst(10) = "Code beheerder"
TLBType(10) = "A"

TLBCode(11) = "A   " + "C003"
TLBTekst(11) = "Naam beheerder"
TLBType(11) = "A"

TLBCode(12) = "N   " + "C004"
TLBTekst(12) = "Telefoonnummer beheerder"
TLBType(12) = "N"

TLBCode(13) = "N   " + "C020"
TLBTekst(13) = "Opening dossier datum"
TLBType(13) = "N"

TLBCode(14) = "A02 " + "C031"
TLBTekst(14) = "Betrokken waarborg 1"
TLBType(14) = "A"

TLBCode(15) = "A03 " + "C032"
TLBTekst(15) = "Betrokken waarborg 2"
TLBType(15) = "A"

TLBCode(16) = "A04 " + "C033"
TLBTekst(16) = "Betrokken waarborg 3"
TLBType(16) = "A"

TLBCode(17) = "A   " + "C100"
TLBTekst(17) = "Tegenpartij naam"
TLBType(17) = "A"

TLBCode(18) = "N   " + "C110"
TLBTekst(18) = "Tegenpartij maatschappij"
TLBType(18) = "N"

TLBCode(19) = "A   " + "C111"
TLBTekst(19) = "Tegenpartij referte"
TLBType(19) = "A"

TLBCode(20) = "A   " + "C400"
TLBTekst(20) = "Expert naam"
TLBType(20) = "A"

TLBCode(21) = "N   " + "C401"
TLBTekst(21) = "Expert aggregatienummer"
TLBType(21) = "N"

TLBCode(22) = "N   " + "C40A"
TLBTekst(22) = "Expert telefoonnummer"
TLBType(22) = "N"

TLBCode(23) = "A   " + "C405"
TLBTekst(23) = "Expert referte"
TLBType(23) = "A"

TLBCode(24) = " 05 " + "C406"
TLBTekst(24) = "Expert type"
TLBType(24) = "N"

TLBCode(25) = " 06 " + "-001"
TLBTekst(25) = "Expertise type"
TLBType(25) = "N"

TLBCode(26) = " 07 " + "C413"
TLBTekst(26) = "Expertise aard"
TLBType(26) = "N"

TLBCode(27) = "N   " + "C410"
TLBTekst(27) = "Datum verzending expert"
TLBType(27) = "N"

TLBCode(28) = "A   " + "C300"
TLBTekst(28) = "Plaats te bezichtigen"
TLBType(28) = "A"

TLBCode(29) = "N   " + "C412"
TLBTekst(29) = "Datum ontvangst expertiseverslag"
TLBType(29) = "N"

TLBCode(30) = "N   " + "C450"
TLBTekst(30) = "Expertise bedrag 1 of forfait"
TLBType(30) = "N"

TLBCode(31) = "N   " + "C451"
TLBTekst(31) = "Expertise bedrag 2"
TLBType(31) = "N"

TLBCode(32) = "N   " + "C453"
TLBTekst(32) = "Totaal verlies"
TLBType(32) = "N"

TLBCode(33) = "A   " + "ZZZZ1"
TLBTekst(33) = "Opmerkingen"
TLBType(33) = "A"

TLBCode(34) = "N   " + "C455"
TLBTekst(34) = "Aantal dagen gebruiksderving"
TLBType(34) = "N"

TLBCode(35) = "A   " + "5020"
TLBTekst(35) = "Plaat voertuig"
TLBType(35) = "A"

TLBCode(36) = "A   " + "5021"
TLBTekst(36) = "Onderstelnummer voertuig"
TLBType(36) = "A"

TLBCode(37) = "A   " + "5000"
TLBTekst(37) = "Voertuig merk/type"
TLBType(37) = "A"

TLBCode(38) = "N   " + "5035"
TLBTekst(38) = "Verzekerde waarde voertuig"
TLBType(38) = "N"

TLBCode(39) = "N08 " + "5400"
TLBTekst(39) = "Type verzekerde waarde"
TLBType(39) = "N"

TLBCode(40) = "N09 " + "5F00"
TLBTekst(40) = "Vrijstelling type 1"
TLBType(40) = "N"

TLBCode(41) = "N09 " + "5F10"
TLBTekst(41) = "Vrijstelling type 2"
TLBType(41) = "N"

TLBCode(42) = "N09 " + "5F20"
TLBTekst(42) = "Vrijstelling type 3"
TLBType(42) = "N"

TLBCode(43) = "A   " + "5F02"
TLBTekst(43) = "Vrijstelling omschrijving"
TLBType(43) = "A"

TLBCode(44) = "N   " + "5F01"
TLBTekst(44) = "Vrijstelling waarde 1"
TLBType(44) = "N"

TLBCode(45) = "N   " + "5F11"
TLBTekst(45) = "Vrijstelling waarde 2"
TLBType(45) = "N"

TLBCode(46) = "N   " + "5F21"
TLBTekst(46) = "Vrijstelling waarde 3"
TLBType(46) = "N"

TLBCode(47) = "A   " + "C6B1"
TLBTekst(47) = "Code uitkering waarborg 1"
TLBType(47) = "A"

TLBCode(48) = "A   " + "C6B2"
TLBTekst(48) = "Uitkering waarborg 1 (omschrijving)"
TLBType(48) = "A"

TLBCode(49) = "N   " + "C6B3"
TLBTekst(49) = "Bedrag uitkering waarborg 1"
TLBType(49) = "N"

TLBCode(50) = "A   " + "C6C1"
TLBTekst(50) = "Code uitkering waarborg 2"
TLBType(50) = "A"

TLBCode(51) = "A   " + "C6C2"
TLBTekst(51) = "Uitkering waarborg 2 (omschrijving)"
TLBType(51) = "A"

TLBCode(52) = "N   " + "C6C3"
TLBTekst(52) = "Bedrag uitkering waarborg 2"
TLBType(52) = "N"

TLBCode(53) = "A   " + "C6D1"
TLBTekst(53) = "Code uitkering waarborg 3"
TLBType(53) = "A"

TLBCode(54) = "A   " + "C6D2"
TLBTekst(54) = "Uitkering waarborg 3 (omschrijving)"
TLBType(54) = "A"

TLBCode(55) = "N   " + "C6D3"
TLBTekst(55) = "Bedrag uitkering waarborg 3"
TLBType(55) = "N"

TLBCode(56) = "N   " + "C029"
TLBTekst(56) = "Datum afsluiting"
TLBType(56) = "N"

TLBCode(57) = ""

End Sub

Function VernieuwMAPIBerichten()

MPIBericht.FetchMsgType = Mid(cbAS2BerichtType.Text, 1, InStr(cbAS2BerichtType.Text, ":") - 1)
MPIBericht.Fetch

lblAantalBerichten.Caption = MPIBericht.MsgCount
lbBerichten.Clear
RichTextBox1.TextRTF = ""
If MPIBericht.MsgCount > 0 Then
    On Error Resume Next
    For COUNT_TO = 0 To MPIBericht.MsgCount - 1
        MPIBericht.MsgIndex = COUNT_TO
        If Err Then Exit For
        lbBerichten.AddItem Format(MPIBericht.MsgDateReceived, "yyyymmdd") & " " & MPIBericht.MsgSubject
    Next
    On Error GoTo 0
    If lbBerichten.ListCount Then
        lbBerichten.ListIndex = lbBerichten.ListCount - 1
        On Error Resume Next
        lbBerichten.SetFocus
    End If
End If

cmdEManueel.Enabled = False
cmdPManueel.Enabled = False
cbPrintNote.Enabled = False
If MPIBericht.MsgCount > 0 Then
    If Mid(cbAS2BerichtType.Text, 1, InStr(cbAS2BerichtType.Text, ":") - 1) = "IPC.AS2.E1" Then
        cmdEManueel.Enabled = True
        cmdEManueel.Default = True
    ElseIf Mid(cbAS2BerichtType.Text, 1, InStr(cbAS2BerichtType.Text, ":") - 1) = "IPM.NOTE.ASSURNET" Then
        cbPrintNote.Enabled = True
        cbPrintNote.Default = True
    Else
        cmdPManueel.Enabled = True
        cmdPManueel.Default = True
    End If
    
End If

End Function

Function refreshMPIBericht() As Boolean


    MPIBericht.FetchMsgType = "IPM"
    MPIBericht.Fetch

    lblAantalBerichten.Caption = MPIBericht.MsgCount
    Me.cbMailBerichten.Clear
    Me.rtbMail.TextRTF = ""
    
    Dim aantalRB As Integer
    Dim cbMailTekstLijn As String
    
    aantalRB = 0
    Me.cbMailBerichten.Clear
    If MPIBericht.MsgCount > 0 Then
        On Error Resume Next
        For COUNT_TO = 0 To MPIBericht.MsgCount - 1
            MPIBericht.MsgIndex = COUNT_TO
            ' TODO
            ' eerst kijken of het mailadres in orde is
            If InStr(Me.MPIBericht.RecipAddress, Me.txtMailBox.Text) Then
                cbMailTekstLijn = "ok RBR-mail:=" & Me.MPIBericht.RecipAddress  'MPIBericht.MsgSubject
                'dan controleren of er attachment is met TB2
                If Me.MPIBericht.AttachmentCount <> 0 Then
                    aantalRB = aantalRB + 1
                    ' ? MPIBericht.AttachmentPathName
                    ' C:\Users\Jos\AppData\Local\Temp\BR298247.txt
                    'me.rtbMail.LoadFile MPIBericht.AttachmentPathName
                    cbMailTekstLijn = cbMailTekstLijn & " Bijlage Gevonden : " & Me.MPIBericht.AttachmentPathName
               Else
                    cbMailTekstLijn = cbMailTekstLijn & " Geen Bijlage Gevonden/Ongeldig: " & Me.MPIBericht.MsgSubject
                End If
            Else
                cbMailTekstLijn = "niet ok voor RBR:=" & Me.MPIBericht.RecipAddress & " : " & MPIBericht.MsgSubject
            End If
    
            If Err Then Exit For
            Me.cbMailBerichten.AddItem cbMailTekstLijn ', Me.cbMailBerichten.ListCount - 1
        Next
        Me.tbTelRetours.Text = Str(aantalRB)
    On Error GoTo 0
    If Me.cbMailBerichten.ListCount Then
        Me.cbMailBerichten.ListIndex = 0
        On Error Resume Next
        Me.cbMailBerichten.SetFocus
    End If
End If

End Function


Private Sub KeuzeInfo_Click(Index As Integer)

    If rsTB2.State = adStateOpen Then
        rsTB2.Close
    End If
    Msg = "SELECT * FROM TB2 WHERE Mij = '" & Left(KeuzeInfo(0).Text, 4) & "'"
    rsTB2.Open Msg, adntDB, adOpenDynamic, adLockOptimistic, adCmdText
    
    lbBerichtenHistoriek.Clear
    RichTextBoxHistoriek.Text = ""
    If rsTB2.RecordCount = 0 Then
    Else
        Do While Not rsTB2.EOF
            lbBerichtenHistoriek.AddItem rsTB2("DatumAanvang").Value & " " & rsTB2("DocType").Value & " " & rsTB2("Polis").Value
            rsTB2.MoveNext
        Loop
        lbBerichtenHistoriek.ListIndex = 0
    End If
    
End Sub

Private Sub lbBerichten_Click()

MPIBericht.MsgIndex = lbBerichten.ListIndex
RichTextBox1.Text = MPIBericht.MsgNoteText

End Sub

Private Sub InstallMij()

Dim T As Integer

KeuzeInfo(0).Clear
T = -1
bGetOrGreater TABLE_SUPPLIERS, 0, "CO"
If Ktrl Then
    Beep
    Exit Sub
Else
    RecordToVeld TABLE_SUPPLIERS
End If
Screen.MousePointer = vbHourglass
If Left(KEY_BUF(TABLE_SUPPLIERS), 2) <> "CO" Then
    Beep
Else
    T = T + 1
    KeuzeInfo(0).AddItem Mid(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 3, 4) + ": " + vBibTekst(TABLE_SUPPLIERS, "#A100 #") + "/" + vSet(vBibTekst(TABLE_SUPPLIERS, "#A400 #"), 8)
    Do
        bNext TABLE_SUPPLIERS
        If Ktrl Or Left(KEY_BUF(TABLE_SUPPLIERS), 2) <> "CO" Then
            Exit Do
        Else
            RecordToVeld TABLE_SUPPLIERS
            T = T + 1
            KeuzeInfo(0).AddItem Mid(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 3, 4) + ": " + vBibTekst(TABLE_SUPPLIERS, "#A100 #") + "/" + vSet(vBibTekst(TABLE_SUPPLIERS, "#A400 #"), 8)
        End If
    Loop
End If
Screen.MousePointer = vbNormal
If KeuzeInfo(0).ListCount Then
    KeuzeInfo(0).ListIndex = 0
End If

End Sub


Private Sub lbBerichtenHistoriek_Click()

    rsTB2.AbsolutePosition = lbBerichtenHistoriek.ListIndex + 1
    RichTextBoxHistoriek.Text = rsTB2("memoTB2").Value
    
End Sub


Private Sub obBorderel_Click()

    gboFILE = "PRENOT.GBO"
    If ScrLeesBestandAlleTekst(gboCPstr, gboDIR & gboFILE) Then
        Me.RichTextBox2.Text = gboCPstr
    Else
        Me.RichTextBox2.Text = ""
    End If
    Label1(4).Caption = gboDIR & gboFILE
    
End Sub


Private Sub opRekeningUittreksel_Click()

    gboFILE = "CURRAC.GBO"
    If ScrLeesBestandAlleTekst(gboCPstr, gboDIR & gboFILE) Then
        Me.RichTextBox2.Text = gboCPstr
    Else
        Me.RichTextBox2.Text = ""
    End If
    Label1(4).Caption = gboDIR & gboFILE
    
End Sub


Private Sub oGBO300_Click()

    gboFILE = "UNIFEP.300"
    If ScrLeesBestandAlleTekst(gboCPstr, gboDIR & gboFILE) Then
        Me.RichTextBox2.Text = gboCPstr
    Else
        Me.RichTextBox2.Text = ""
    End If
    Label1(4).Caption = gboDIR & gboFILE

End Sub


Private Function DoTheGBO() As Boolean
        
Dim MAPIString As String
Dim StringXGH As String
Dim StringXEH As String
Dim StringXET As String
Dim StringXGT As String
Dim MijNummer As String * 4

Dim Xarray() As String
Dim XGHarray() As String
Dim XEHarray() As String
Dim intCrLfOpkuis As Integer


MAPIString = RichTextBox1.Text

'nazien of er geen CRLF's zijn ! AXA en GBO...
    intCrLfOpkuis = 0
    Do While InStr(MAPIString, vbCrLf)
        MAPIString = Replace(MAPIString, vbCrLf, "")
        intCrLfOpkuis = intCrLfOpkuis + 1
    Loop
    If intCrLfOpkuis Then
        'MsgBox Str(intCrLfOpkuis) + " x crlf verwijderd", vbInformation
    End If
    
Xarray = Split(MAPIString, "'")
XGHarray = Split(Xarray(0) + "+", "+")
XEHarray = Split(Xarray(1) + "+", "+")

Dim XEHcode As String
Dim XEHcodeHier As String

    XEHcode = Format(XGHarray(2), "0000")
    'XEHcodeHier = Format(XGHarray(2), "0000")
    MijNummer = XEHcode 'XGHarray(3)

If InStr(MAPIString, "XEH+01") Then
    'XEH+01 = NIEUWE ZAAK OF WIJZIGING GEVOLGD EVENTUEEL DOOR XEH+03
    MsgBox "Wijziging of nieuwe zaak"
    EdiFactTERMIJN MAPIString
    XEHcodeHier = XEHarray(1) '"01"
    strA000RB = XEHarray(4)
    If Trim(strA000RB) = "" Then
        Msg = Mid(MAPIString, InStr(MAPIString, "RFF+001:") + 8)
        If Trim(Msg) = "" Then
            Stop
        Else
            strA000RB = Mid(Msg, 1, InStr(Msg, "'") - 1)
        End If
        If Trim(strA000RB) = "" Then
            Stop
        End If
    Else
        SnelHelpPrint "Uitwisseling voor polis " & strA000RB & " verzekeraar RETOURBLOK " & XGHarray(2) & " wordt bewaard.", False
    End If
    GoSub VerwijderENslaOP
    
ElseIf InStr(MAPIString, "XEH+02") Then
    'Stop
    MsgBox "Opening schadedossier " & XGHarray(2), vbInformation
    'SchadeBeheer XEHarray(3), MAPIString, XGHarray(2)
    XEHcodeHier = "02"
    If Trim(strA000RB) = "" Then
        Msg = Mid(MAPIString, InStr(MAPIString, "RFF+001:") + 8)
        If Trim(Msg) = "" Then
            Stop
        Else
            strA000RB = Mid(Msg, 1, InStr(Msg, "'") - 1)
        End If
        If Trim(strA000RB) = "" Then
            Stop
        End If
    Else
        SnelHelpPrint "Uitwisseling voor polis " & strA000RB & " verzekeraar RETOURBLOK " & XGHarray(2) & " wordt bewaard.", False
    End If
    GoSub VerwijderENslaOP
ElseIf InStr(MAPIString, "XEH+03") Then
    'XEH+03 = KWIJTING KOMT BIJ BIJVOEGSELS ACHTERAAN DUS HIER ZIJN HET HERNIEUWINGEN
    MsgBox "Uitwisseling Termijn " & XGHarray(2), vbInformation
    XEHcodeHier = "03"
    If Me.obVerwerking(0).Value = True Then
        EdiFactTERMIJN MAPIString
    End If
    GoSub VerwijderENslaOP
ElseIf InStr(MAPIString, "XEH+07") Then
    'XEH+07 = REKENINGUITTREKSELS  ?
    MsgBox "Rekeninguittreksel " & XGHarray(2), vbInformation
    XEHcodeHier = "07"
    If Me.obVerwerking(0).Value = True Then
        EdiFactREKENINGUITTREKSEL MAPIString
    End If
    GoSub VerwijderENslaOP
Else
    MsgBox "Nog niets voorzien voor verzoek " & XEHarray(3) & " maatschappij " & XGHarray(2)
    Exit Function
    GoSub VerwijderENslaOP
End If
Exit Function

VerwijderENslaOP:
    TLB_RECORD(TABLE_VARIOUS) = ""
    vBib TABLE_VARIOUS, vSet("L" + "CO" + Format(XGHarray(2), "0000"), 13), "v004"
    vBib TABLE_VARIOUS, Format(Left(XGHarray(2), 4), "0000"), "A010"
    vBib TABLE_VARIOUS, "E1", "v400"
    vBib TABLE_VARIOUS, vSet("30" + Format(Left(XGHarray(2), 4), "0000") + "E1", 20), "v005"
    vBib TABLE_VARIOUS, RichTextBox1.Text, "v132"
    'rsmar(flallerlei)("A000")= djfkddk
    'rsMAR(TABLE_VARIOUS).Update
    
    rsTB2.AddNew
    rsTB2("Mij").Value = Format(Left(XGHarray(2), 4), "0000")
    rsTB2("memoTB2").Value = MAPIString
    rsTB2("DocType").Value = XEHcodeHier
    rsTB2("Polis").Value = strA000RB
    'nu nog de datum op basis van doctype
    If XEHcodeHier = "01" Then 'contract!
    ElseIf XEHcodeHier = "02" Then 'schadegeval
        posSTR = InStr(MAPIString, "DTM+048:")
        If posSTR = 0 Then
            MsgBox "Stop"
        Else
            Msg = Mid(MAPIString, posSTR + 8)
            Msg = Left(Msg, 8)
        End If
        'rsTB2("DatumAanvang").Value = CVDate(Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) & "/" & Mid(Msg, 5, 4))
    ElseIf XEHcodeHier = "07" Then 'rekeninguittreksel?
        rsTB2("Polis").Value = "" 'zeker geen polisnummer dan!
        posSTR = InStr(MAPIString, "DTM+069:")
        If posSTR = 0 Then
            MsgBox "Stop"
        Else
            Msg = Mid(MAPIString, posSTR + 8)
            Msg = Left(Msg, 8)
        End If
        If Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) = "29/02" Then
            'best 28/02 van maken
            rsTB2("DatumAanvang").Value = CVDate("28/02" & "/" & Mid(Msg, 5, 4))
            
        Else
            rsTB2("DatumAanvang").Value = CVDate(Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 2) & "/" & Mid(Msg, 5, 4))
        End If
                
    ElseIf XEHcodeHier = "03" Then 'hernieuwingen ?
        rsTB2("Polis").Value = "" 'zeker geen polisnummer dan!
        posSTR = InStr(MAPIString, "DTM+005:")
        If posSTR = 0 Then
            MsgBox "Stop"
        Else
            Msg = Mid(MAPIString, posSTR + 8)
            Msg = Left(Msg, 6)
        End If
        rsTB2("DatumAanvang").Value = CVDate("01" & "/" & Mid(Msg, 1, 2) & "/" & Mid(Msg, 3, 4))
    Else
        MsgBox "Stop"
    End If
    If Me.obVerwerking(0).Value = True Then
        rsTB2("status") = "1"
    Else
        rsTB2("status") = "0"
    End If
    rsTB2.Update
    bInsert TABLE_VARIOUS, 1
    'MsgBox "Stop"
    If Ktrl Then
        MsgBox Error
    Else
        RichTextBox1.Text = ""
    End If
Return

End Function

Private Function DoTheGBO300() As Boolean
        
Dim MAPIString As String
Dim MijNummer As String * 4

Dim RuilKaart As String
Dim strLijn As String
Dim strFROM As String
Dim strTO As String
Dim strAGENTNR As String
Dim strVERRICHTING As String
        
    RuilKaart = Me.RichTextBox2.Text
    Do While RuilKaart <> ""
        strLijn = Left(RuilKaart, InStr(RuilKaart, vbCrLf) - 1)
        If Len(strLijn) <> 300 Then
            MsgBox "Stop"
        End If
        strFROM = Mid(strLijn, 1, 4)
        strTO = Mid(strLijn, 5, 4)
        strAGENTNR = Trim$(Mid(strLijn, 9, 10))
        strVERRICHTING = Mid(strLijn, 19, 2)
        Select Case strVERRICHTING
            Case "11" 'kwijting inning makelaar
                Stop
                
                
                
                
                
                
                
                
            Case "31" 'kwijting inning maatschappij"
                'overslaan voor mar!
            Case "20" To "29" 'verrichtingen in uittreksel
                'voorlopig overslaan.  Bestaan tevens uit slechts één record
        End Select
        RuilKaart = Mid(RuilKaart, InStr(RuilKaart, vbCrLf) + 2)
    Loop

    MsgBox "Einde bestand.  Alle records verwerkt", vbInformation

End Function

Private Sub SSTab1_Click(PreviousTab As Integer)


    Select Case SSTab1.Tab
        Case 0
            'MsgBox "Stop"
        Case 1
            If strTELEBIBIO = "" Then
            Else
                RichTextBoxHistoriek.Text = strTELEBIBIO
            End If
            
        Case 2
            'MsgBox "Stop"
        Case 3
            MsgBox "Stop"
        
        Case 4
            'MsgBox "Stop"
        
    End Select
    
End Sub

Private Sub cbFetch_Click()
    
    On Error Resume Next
    
ProbeerNogEens:
    DoEvents
    Me.MPIBericht.Fetch
    If Err = 32053 Then
        GoSub mailLogon
        GoTo ProbeerNogEens
    Else
        Me.refreshMPIBericht
        Me.cbMailBerichten.ListIndex = 0
    End If
    DoEvents
    Exit Sub
    
mailLogon:
    Err = 0
    MPISessie.DownLoadMail = True
    MPISessie.SignOn
    If Err Then
        MsgBox "Mail niet beschikbaar" & vbCrLf & Error, vbInformation
        Me.cbMailBerichten.Enabled = False
    Else
        MPIBericht.SessionID = MPISessie.SessionID
    End If
Return

End Sub
