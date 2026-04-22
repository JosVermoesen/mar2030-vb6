VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Xlog 
   AutoRedraw      =   -1  'True
   Caption         =   "Log"
   ClientHeight    =   5325
   ClientLeft      =   2880
   ClientTop       =   660
   ClientWidth     =   6690
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   21100
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5325
   ScaleWidth      =   6690
   Begin TabDlg.SSTab SSTab1 
      Height          =   4815
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   8493
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Default"
      TabPicture(0)   =   "0xlog.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "X"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Afbeelding"
      TabPicture(1)   =   "0xlog.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "OLE1"
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(4)=   "Command1"
      Tab(1).Control(5)=   "Command2"
      Tab(1).Control(6)=   "Command3"
      Tab(1).Control(7)=   "Command4"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton Command4 
         Caption         =   "Pdf Opslaan"
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
         Height          =   375
         Left            =   -74880
         TabIndex        =   11
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Tif Bewerken"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   10
         Top             =   1560
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Pdf Opslaan"
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
         Height          =   375
         Left            =   -74880
         TabIndex        =   8
         Top             =   600
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Pdf Bewerken"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   7
         Top             =   120
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid X 
         Height          =   4485
         HelpContextID   =   21100
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   7911
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   12648384
         FocusRect       =   2
         HighLight       =   2
         ScrollBars      =   2
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label3 
         Caption         =   $"0xlog.frx":0038
         Height          =   3015
         Left            =   -70800
         TabIndex        =   14
         Top             =   360
         Width           =   1815
      End
      Begin VB.OLE OLE1 
         Class           =   "AcroExch.Document.7"
         Height          =   3255
         Left            =   -73200
         SizeMode        =   1  'Stretch
         TabIndex        =   13
         Top             =   120
         Width           =   2295
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   690
         Left            =   -75000
         TabIndex        =   12
         Top             =   3600
         Width           =   6375
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   1080
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdDetailJournaal 
      Caption         =   "Detail &Journaal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4500
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4920
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.CommandButton cbAfbeelding 
      Caption         =   "Af&drukken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3060
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1400
   End
   Begin VB.CommandButton Afsluiten 
      Caption         =   "Ok"
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
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   4920
      Width           =   1400
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "Terugkeren"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1620
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1400
   End
   Begin VB.CommandButton WijzigenLijn 
      Caption         =   "Wijzigen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4500
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1400
   End
   Begin VB.Menu Bewerken 
      Caption         =   "&Bewerken"
      Begin VB.Menu Kopie 
         Caption         =   "&Kopiëren"
         Index           =   0
      End
      Begin VB.Menu Kopie 
         Caption         =   "&Selecteren mogelijk"
         Index           =   1
      End
      Begin VB.Menu Kopie 
         Caption         =   "&Grafische afdruk"
         Index           =   2
      End
      Begin VB.Menu Kopie 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu Kopie 
         Caption         =   "MsDos Editor Laserprinter Bestand (&HP Stuurcodes)"
         Index           =   4
      End
      Begin VB.Menu Kopie 
         Caption         =   "MsDos Editor Kettingprinter Bestand (&IBM Stuurcodes)"
         Index           =   5
      End
      Begin VB.Menu Kopie 
         Caption         =   "MsDos Editor Puur Tekstbestand "
         Index           =   6
      End
      Begin VB.Menu Kopie 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu Kopie 
         Caption         =   "Opslaan met scheidingstekens"
         Index           =   8
      End
      Begin VB.Menu Kopie 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu Kopie 
         Caption         =   "Standaard &Venstergrootte"
         Index           =   10
      End
   End
End
Attribute VB_Name = "Xlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

'Dim miDoc As MODIctl.Document
Dim bstPDFofTIF As String

Dim MaxPlijn As Integer
Dim ATLijn          As Integer

Dim OptieTxt               As String
Dim TempoTxt               As String
Dim CrText                 As String
Dim CrText2                As String

Sub BewaarAls()
Dim Tel As Integer
Dim FlTempBewaar As Integer
Dim DeString As String
Dim TabPos As Integer
Dim LijstSep As String

On Local Error Resume Next
Err = 0
LijstSep = LaadTekst("c:\windows\win.ini;intl", "sList")
If Err Or LijstSep = "" Then
    LijstSep = ";"
    'MsgBox "Kan lijstseparator in internationale instellingen niet vinden.  ';' wordt automatisch genomen...  Kontroleer in ieder geval de systeeminstellingen !"
End If
Screen.MousePointer = vbHourglass

On Local Error GoTo CancelError
Mim.Teken.filename = ""
Mim.Teken.CancelError = True
Mim.Teken.Filter = "Alle bestanden (*.*)|*.*"
Mim.Teken.ShowSave

FlTempBewaar = FreeFile
Open Mim.Teken.filename For Output As FlTempBewaar

For Tel = 1 To X.Rows - 1
    X.Row = Tel
    X.ColSel = X.Cols - 1
    DeString = X.Clip
    TabPos = InStr(DeString, vbTab)
    Do While TabPos > 0
        Mid(DeString, TabPos, 1) = LijstSep
        TabPos = InStr(DeString, vbTab)
    Loop
    Print #FlTempBewaar, DeString; vbCrLf;
Next
Close FlTempBewaar

CancelError:
Screen.MousePointer = vbNormal

End Sub


Private Sub Afsluiten_Click()
Dim T               As Integer
Dim xPos            As Integer
Dim xPos2           As Integer

X.Col = 0
If X.text = "" Then
    If Left(Xlog.Caption, 6) = "Schade" Then
        XLogKey = "Nieuw"
    End If
Else
    XLogKey = X.text + vbCrLf
    X.Col = 1
    XLogKey = XLogKey + X.text
End If
Xlog.Hide

End Sub

Private Sub Annuleren_Click()

XLogKey = ""
GridText = ""
Xlog.WindowState = 0
Xlog.Hide

End Sub


Private Sub cbAfbeelding_Click()

Dim OnthoudCol As Integer
Dim OnthoudRow As Integer
Dim telrow As Long

Msg = "Grafische afdruk van het venster !  Bent U zeker ?" + vbCrLf + vbCrLf
Msg = Msg + "Kies 'Nee' voor gewone afdruk van alle ingevulde tekstlijnen." + vbCrLf
Msg = Msg + "Hiermee voldoet U aan uw GPRD verplichtingen indien uw partij zijn/haar gegevens opvraagt"

KtrlBox = MsgBox(Msg, vbYesNoCancel + vbDefaultButton3)
Select Case KtrlBox
    Case vbYes
        'MsgBox "Grafische schermafdruk is tijdelijk niet mogelijk"
        On Error Resume Next
        PrintForm
    Case vbNo
    MousePointer = 11
    OnthoudCol = X.Col
    OnthoudRow = X.Row
    
    PAGE_COUNTER = 0
    Set Printer = Printers(LijstPrinterNr)
    On Error Resume Next
    Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
    Printer.FontName = "Courier New"
    Printer.FontSize = 12
    Printer.FontBold = True
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print UCase(Mid(Mim.Caption, InStr(Mim.Caption, "[")));
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf;
        Printer.Print "marINTEGRAAL NT : ";
    Printer.FontBold = True
        Printer.Print Xlog.Caption;
    Printer.FontBold = False
        Printer.Print vbCrLf;
        Printer.Print Left(FULL_LINE, 80); vbCrLf;
    
    For telrow = 1 To X.Rows - 1
        X.Row = telrow
        X.Col = 2
        If RTrim$(X.text) = "" Then
        Else
            X.Col = 1
            Printer.FontBold = False
            Printer.Print Tab(1); X.text;
            Printer.FontBold = True
            X.Col = 2
            If Len(X.text) > 40 Then
                Printer.Print vbCrLf;
                Printer.Print Tab(1); X.text;
            Else
                Printer.Print Tab(40); X.text;
            End If
            Printer.Print vbCrLf;
        End If
    Next
    Printer.FontBold = False
        Printer.Print vbCrLf; vbCrLf; "Datum en uur vandaag             : ";
    Printer.FontBold = True
        Printer.Print Format(Now, "dd/mm/yyyy hh:mm:ss")
    On Local Error Resume Next
    Printer.NewPage
    Printer.EndDoc
    X.Col = OnthoudCol
    X.Row = OnthoudRow
    MousePointer = 0
    
End Select

End Sub



Private Sub CmdDetailJournaal_Click()
Dim A As String
Dim DeString As String

X.Col = 0
DeString = X.text

Unload InfoScherm
Load InfoScherm
InfoScherm.Caption = InfoScherm.Caption + ", Journaaldetail voor dokument : " + DeString
InfoScherm.X.Rows = 2
InfoScherm.X.Cols = 8
InfoScherm.X.Col = 0
InfoScherm.X.Row = 0
InfoScherm.X.text = "Datum #v066"
InfoScherm.X.ColAlignment(0) = flexAlignLeftCenter
InfoScherm.X.Col = 1
InfoScherm.X.text = "Rekening #v019"
InfoScherm.X.ColAlignment(1) = flexAlignLeftCenter
InfoScherm.X.Col = 2
InfoScherm.X.text = "Naam #v020"
InfoScherm.X.ColAlignment(2) = flexAlignLeftCenter
InfoScherm.X.Col = 3
InfoScherm.X.text = "Bedrag #v068"
InfoScherm.X.ColAlignment(3) = flexAlignGeneral
InfoScherm.X.Col = 4
InfoScherm.X.text = "Boekingsomschrijving #v067"
InfoScherm.X.ColAlignment(4) = flexAlignLeftCenter
InfoScherm.X.Col = 5
InfoScherm.X.text = "Fin. Stuk #v038"
InfoScherm.X.ColAlignment(5) = flexAlignLeftCenter
InfoScherm.X.Col = 6
InfoScherm.X.text = "TegenRek. #v069"
InfoScherm.X.ColAlignment(6) = flexAlignLeftCenter
InfoScherm.X.Col = 7
InfoScherm.X.text = "vsfRecord"


bGet TABLE_JOURNAL, 1, vSet(DeString, 11)
If Ktrl Then
    Beep
    MsgBox "Geen journaallijnen voor " + DeString
    Exit Sub
Else
    Screen.MousePointer = vbHourglass
    GoSub VolgendeLijn
    Do
        bNext TABLE_JOURNAL
        If Ktrl Or KEY_BUF(TABLE_JOURNAL) <> DeString Then
            Exit Do
        Else
            GoSub VolgendeLijn
        End If
    Loop
End If
Screen.MousePointer = vbNormal

OpHetScherm:
InfoScherm.X.Row = 1
InfoScherm.X.Col = 0

InfoScherm.X.ColWidth(0) = 1005
InfoScherm.X.ColWidth(1) = 765
InfoScherm.X.ColWidth(2) = 2160
InfoScherm.X.ColWidth(3) = 1000
InfoScherm.X.ColWidth(4) = 1950
InfoScherm.X.ColWidth(5) = 855
InfoScherm.X.ColWidth(6) = 870
InfoScherm.X.ColWidth(7) = 6000
InfoScherm.X.ColAlignment(0) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(1) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(2) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(3) = flexAlignGeneral
InfoScherm.X.ColAlignment(4) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(5) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(6) = flexAlignLeftCenter
InfoScherm.X.ColAlignment(7) = flexAlignLeftCenter

XLogShow:
InfoScherm.Show 1
Unload InfoScherm
Exit Sub

VolgendeLijn:
RecordToVeld TABLE_JOURNAL
A = DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #")) & vbTab
A = A & vBibTekst(TABLE_JOURNAL, "#v019 #") & vbTab
bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 7)
If Ktrl Then
    A = A & "//" & vbTab
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    A = A & vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #") & vbTab
End If
A = A & Format(rsMAR(TABLE_JOURNAL)("dece068"), "#,##0.00") & vbTab
A = A & rsMAR(TABLE_JOURNAL)("v067") & vbTab
A = A & rsMAR(TABLE_JOURNAL)("v038") & vbTab
A = A & rsMAR(TABLE_JOURNAL)("v069") & vbTab
A = A & TLB_RECORD(TABLE_JOURNAL)
InfoScherm.X.AddItem A, InfoScherm.X.Rows - 1
Return

End Sub


Private Sub Command1_Click()

    Command3.Enabled = False
    Command4.Enabled = False

    Dim path As String * 260
    Dim ret As Integer
    'Programmalokatie ?
    ret = SHGetFolderPath(0, CSIDL_PROGRAM_FILES, 0, 0, path)
    Dim wProgramFiles As String
    wProgramFiles = Left(path, InStr(path, Chr(0)) - 1)
    
    Dim flHier As Integer
    Me.Command2.Enabled = False
    If Me.Command1.Caption = "Pdf Openen" Then
        If LaadTekst("dnnInstellingen", "Mario") = "" Then
            MsgBox "Nieuwe PC of nog geen instellingen voor PDF Postvak In.  Aanbevolen in te te stellen a.u.b. via submenu DotNetNuke.", vbInformation
            Mim.Teken.InitDir = LOCATION_COMPANYDATA
        Else
            Mim.Teken.InitDir = LaadTekst("dnnInstellingen", "Mario")
        End If
        Mim.Teken.Filter = "Acrobat PDF bestand (*.pdf)|*.pdf"
        Mim.Teken.filename = ""
        Mim.Teken.ShowOpen
        bstPDFofTIF = LCase(Mim.Teken.filename)
        If bstPDFofTIF = "" Then
            Exit Sub
        ElseIf Right(bstPDFofTIF, 3) <> "pdf" Then
            MsgBox "Uitsluitend Pdf bestanden selecteren a.u.b.", vbExclamation
            'Exit Sub
        End If
        bstPDFofTIF = Mim.Teken.filename
        Label2.Caption = bstPDFofTIF
        Screen.MousePointer = vbHourglass
        On Local Error Resume Next
        OLE1.CreateEmbed bstPDFofTIF
        Screen.MousePointer = vbNormal
                
        'MiDocView1.filename = bstAfbeeldingHier
        Me.Command2.Enabled = True
        'Me.cbAfdruk.Enabled = True
    ElseIf Me.Command1.Caption = "Pdf Wijzigen" Then
        'MiDocView1.filename = ""
        DoEvents
        Command2_Click
                
    ElseIf Me.Command1.Caption = "Pdf Vernieuwen" Then
        'MiDocView1.filename = bstAfbeeldingHier
        Me.Command2.Enabled = True
        'Me.cbAfdruk.Enabled = True
        Me.Command1.Caption = "Pdf Wijzigen"
    End If
    
End Sub

Private Sub Command2_Click()

    FileToBlob rsMAR(SharedScanFl)("bstBLOB37"), bstPDFofTIF
    rsMAR(SharedScanFl)("bstndNaam37") = bstPDFofTIF
    rsMAR(SharedScanFl)("typeZending37") = Mid(bstPDFofTIF, InStr(bstPDFofTIF, ".") + 1)
    rsMAR(SharedScanFl).Update
    Me.SSTab1.Tab = 0
    
End Sub


Private Sub Command3_Click()

'3 en 4 !!
Command1.Enabled = False
Command2.Enabled = False
    
End Sub




Private Sub Form_Load()

LoadFormProperties Me

'Set miDoc = New MODIctl.Document

End Sub

Private Sub Form_LostFocus()

'Stop

End Sub

Private Sub Form_Resize()

On Local Error Resume Next

Me.SSTab1.Top = 0
Me.SSTab1.Left = 0
Me.SSTab1.Width = Xlog.Width - 100
Me.SSTab1.Height = Xlog.Height - 1300

X.Width = Me.SSTab1.Width - 100
X.Height = Me.SSTab1.Height - 400

Afsluiten.Top = Xlog.Height - 1200
Annuleren.Top = Xlog.Height - 1200
cbAfbeelding.Top = Xlog.Height - 1200
WijzigenLijn.Top = Xlog.Height - 1200
CmdDetailJournaal.Top = Xlog.Height - 1200

'X.Top = 0
'X.Left = 0
'X.Width = Xlog.Width - 100
'X.Height = Xlog.Height - 1300

'Afsluiten.Top = Xlog.Height - 1200
'Annuleren.Top = Xlog.Height - 1200
'CmdAfdrukken.Top = Xlog.Height - 1200
'WijzigenLijn.Top = Xlog.Height - 1200
'CmdDetailJournaal.Top = Xlog.Height - 1200

End Sub

Private Function QuickHelp(InfoString As String) As String
    
'Select Case Left(TELEBIB_CODE(X.Row - 1), 1)
Select Case Left(InfoString, 1)
    Case "1"
        QuickHelp = "Naam of adres"
    Case "2"
        QuickHelp = "Beschrijving, tekst of omschrijving"
    Case "3"
        QuickHelp = "Een Bedrag in + of -"
    Case "4"
        QuickHelp = "Een hoeveelheid (+)"
    Case "5"
        QuickHelp = "Kode (1 of meer tekens)"
    Case "6"
        QuickHelp = "Index (000.00)"
    Case "7"
        QuickHelp = "Referentie"
    Case "8"
        QuickHelp = "Percentage (max. 999)"
    Case "9"
        QuickHelp = "Datum (DDMMEEJJ)"
    Case "A"
        QuickHelp = "Communicatiekanalen (telefoon, fax...)"
    Case "B"
        QuickHelp = "Financiële rekening (xxx-xxxxxxx-xx)"
    Case "b"
        QuickHelp = "Btw Nummer of Nationaal nummer (xxx.xxx.xxx)"
    Case "c", "d"
        QuickHelp = "Geldige bestandsnaam a.u.b"
    Case "z"
        QuickHelp = "Volledige datum als sleutel"
End Select

If Mid(InfoString, 2, 2) <> "  " Then
    SnelHelpPrint " [Ctrl] voor Venster keuzeopties)", BL_LOGGING
End If

End Function




Private Sub Form_Unload(Cancel As Integer)
Dim X As Boolean

X = SaveFormProperties(Me)
'Me.MiDocView1.filename = ""

End Sub

Private Sub Kopie_Click(Index As Integer)

On Error Resume Next
Dim T As Integer
Dim xPosX As Integer
Dim xPosY As Integer
Dim flHier As Integer
Dim TelLijnHier As Variant: TelLijnHier = 72
Dim PuurTekst As Boolean

PuurTekst = False
Select Case Index
    Case 0
        If X.Clip = "" Then
            MsgBox "Eerst selecteren a.u.b. !"
            Exit Sub
        Else
            On Local Error Resume Next
            Err = 0
            Clipboard.Clear
            Clipboard.SetText X.Clip
        End If
        If Err Then
            MsgBox "Kopiëren naar het klembord was onvolledig.  Max. 64000 tekens aan tekst kan geknipt worden.  Uw toestel beschikt mogelijk over onvoldoende geheugen.  Sluit overbodige toepassingen en probeer eventueel opnieuw."
            Exit Sub
        End If
    
    Case 1
        Kopie(1).Checked = Not Kopie(1).Checked
        
    Case 2
        PrintForm
        
    Case 4
        If Dir("printbst.txt") Then
            Kill "printbst.txt"
        End If
        FileCopy App.path & "\Def\hpj.def", "printbst.txt"
        GoSub StelLijstSamen
        
    Case 5
        If Dir("printbst.txt") Then
            Kill "printbst.txt"
        End If
        FileCopy App.path & "\Def\ibm.def", "printbst.txt"
        GoSub StelLijstSamen
        
    Case 6
        If Dir("printbst.txt") Then
            Kill "printbst.txt"
        End If
        PuurTekst = True
        GoSub StelLijstSamen
        
    Case 8
        BewaarAls
        
    Case 10
        Xlog.WindowState = 0
        Xlog.Height = 6465
        Xlog.Width = 9315
        
End Select
Exit Sub

StelLijstSamen:
TelLijnHier = InputBox("Aantal lijnen per blad", "Paginalengte", TelLijnHier)
If TelLijnHier = "" Then Return
                
PAGE_COUNTER = 0
flHier = FreeFile
Open "printbst.txt" For Append As flHier
GoSub PrintDeTITEL
For xPosX = 1 To Xlog.X.Rows - 1
    'controle voor NieuwBladTitelPrint
    For xPosY = 0 To Xlog.X.Cols - 1
        Print #flHier, Tab(REPORT_TAB(xPosY)); X.TextMatrix(xPosX, xPosY);
    Next
    COUNT_TO = COUNT_TO + 1: Print #flHier, vbCrLf;
    If COUNT_TO >= TelLijnHier Then
        If PuurTekst = False Then
            Print #flHier, Chr(12);
        End If
        GoSub PrintDeTITEL
    End If
Next
Close flHier
Shell "edit.com printbst.txt", vbNormalFocus
Return

PrintDeTITEL:
COUNT_TO = 12
PAGE_COUNTER = PAGE_COUNTER + 1
Print #flHier, Tab(1); Mid(Mim.Caption, InStr(Mim.Caption, "[")) & " DosBoxPrint"; Tab(108); "Pagina : "; Dec$((PAGE_COUNTER), "##########"); vbCrLf;
Print #flHier, Tab(108); "Datum  : "; MIM_GLOBAL_DATE; vbCrLf; vbCrLf;
Print #flHier, Tab(1); Xlog.Caption; vbCrLf; String(128, "-"); vbCrLf;

For T = 0 To Xlog.X.Rows - 1
    Print #flHier, Tab(REPORT_TAB(T)); X.TextMatrix(0, T);
Next
Print #flHier, vbCrLf; String(128, "-"); vbCrLf; vbCrLf;
Return

End Sub











Private Sub SSTab1_Click(PreviousTab As Integer)

    Dim strTypeHier As String
    
    If bstPDFofTIF = "" Then
        Me.Label1.Caption = X.TextMatrix(X.Row, 0)
        Me.Command1.Enabled = False
        Me.Command3.Enabled = False
    Else
        Me.Label1.Caption = X.TextMatrix(X.Row, 0)
        Me.Label2.Caption = bstPDFofTIF
        If Mid(bstPDFofTIF, InStr(bstPDFofTIF, ".") + 1) = "pdf" Then
            Me.Command1.Enabled = True
            Me.Command3.Enabled = False
        Else
            Me.Command3.Enabled = True
            Me.Command1.Enabled = False
        End If
    End If
    Select Case SSTab1.Tab
        Case 0
            'MiDocView1.filename = ""
            bstPDFofTIF = ""
            SSTab1.TabCaption(1) = "- Geen Bijlage"
            Me.Label1.Caption = "- leeg -"
            Me.Label2.Caption = ""
            Xlog.cbAfbeelding.FontBold = False
            
        Case 1
            If Left(SSTab1.TabCaption(1), 1) = "+" Then
                strTypeHier = rsMAR(SharedScanFl)("typeZending37")
                bstPDFofTIF = LaadTekst("dnnInstellingen", "Mario") & "\ioDocument." & strTypeHier
                Me.Label2.Caption = bstPDFofTIF
                Screen.MousePointer = vbHourglass
                If Dir(bstPDFofTIF) <> "" Then Kill bstPDFofTIF
                BlobToFile rsMAR(SharedScanFl)("bstBLOB37"), bstPDFofTIF
                DoEvents
                Screen.MousePointer = vbHourglass
                On Local Error Resume Next
                'OLE1.CreateEmbed bstPDFofTIF
                Screen.MousePointer = vbNormal
                'MiDocView1.filename = bstAfbeeldingHier
                Screen.MousePointer = vbHourglass
                On Local Error Resume Next
                OLE1.CreateEmbed bstPDFofTIF
                Screen.MousePointer = vbNormal
                Me.Command1.Caption = "Pdf Wijzigen": Me.Command1.Enabled = True: Me.Command2.Enabled = False
                Me.Command3.Caption = "Tif Wijzigen": Me.Command3.Enabled = True: Me.Command4.Enabled = False
                Label3.Visible = True
            Else
                'het is nieuw
                OLE1.SourceDoc = ""
                OLE1.Refresh
                Me.Command1.Caption = "Pdf Openen": Me.Command1.Enabled = True: Me.Command2.Enabled = False
                Me.Command3.Caption = "Tif Openen": Me.Command3.Enabled = True: Me.Command4.Enabled = False
                Label3.Visible = False
            End If
        Case Else
    End Select

End Sub


Private Sub WijzigenLijn_Click()

On Local Error Resume Next

If Mid(TELEBIB_CODE(X.Row - 1), 2, 2) <> "  " And Mid(TELEBIB_CODE(X.Row - 1), 1, 1) <> "@" Then
    X_KeyDown 17, 0
    If X.Row < X.Rows - 1 Then
        X.Row = X.Row + 1
        If X.Row > 6 Then
            X.TopRow = X.Row - 5
        End If
    End If
    X.SetFocus
Else
    Msg = ""
    If Mid(TELEBIB_CODE(X.Row - 1), 10, 1) = "-" Then
        Msg = "Deze informatie kan niet gewijzigd worden..."
        GridText = "Edit No"
    ElseIf Mid(TELEBIB_CODE(X.Row - 1), 10, 1) = "x" Then
        Msg = "Deze informatie kan niet gewijzigd worden..."
        GridText = "Edit No"
    Else
        GridText = "Edit Yes"
    End If
    If Mid(TELEBIB_CODE(X.Row - 1), 1, 1) = "@" Then
        Msg = Mid(TELEBIB_CODE(X.Row - 1), 1, 3)
    Else
        Msg = Msg + QuickHelp(Left(TELEBIB_CODE(X.Row - 1), 3))
    End If
    X.Col = 2
    ATLijn = Val(Mid(TELEBIB_CODE(X.Row - 1), 10, 1))
    If Mid(TELEBIB_CODE(X.Row - 1), 10, 1) = "x" Then
        If Not IsNull(rsMAR(Val(Mid(X.text, 1, 2)))(Mid(TELEBIB_CODE(X.Row - 1), 5, 4))) Then
            Load FormReactionsDialog
            FormReactionsDialog.TextBoxReactions.text = rsMAR(Val(Mid(X.text, 1, 2)))(Mid(TELEBIB_CODE(X.Row - 1), 5, 4))
            If Val(Mid(X.text, 1, 2)) <= 2 Then
                FormReactionsDialog.Caption = "Ondersteunde documenten"
            Else
                FormReactionsDialog.Caption = "Bevestigingen en reacties"
            End If
            FormReactionsDialog.Show 1
        Else
            MsgBox "Geen gegevens geregistreerd", vbInformation
        End If
    Else
        X.text = vsfInputBox$(Msg, RTrim$(TELEBIB_TEXT(X.Row - 1)), (X.text), "")
        If X.Row < X.Rows - 1 Then
            X.Row = X.Row + 1
            If X.Row > 6 Then
                X.TopRow = X.Row - 5
            End If
        End If
        X.SetFocus
    End If
End If

End Sub

Private Sub X_Click()

On Error Resume Next

If Kopie(1).Checked = True Then Exit Sub
With X
    .Col = 2
    .RowSel = .Row
    .ColSel = .Cols - 1
End With

End Sub

Private Sub X_DblClick()

If Afsluiten.TabStop Then
    WijzigenLijn_Click
Else
    Afsluiten_Click
End If

End Sub




Private Sub X_KeyDown(KeyCode As Integer, Shift As Integer)
Dim DummyText As String
Dim IDXPos As Integer
Dim BoxType As Integer
Dim MBNummer As String

If KeyCode = 17 Then
Else
    Exit Sub
End If
X.Col = X.Cols - 1
DummyText = RTrim$(X.text)
If Left(Caption, 3) = "Log" Then
    Select Case Mid(TELEBIB_CODE(X.Row - 1), 2, 2)
        Case "K ", "L ", "LC", "R ", "R3", "R4", "R6", "R7"
            If KeyCode <> 17 Then Exit Sub
            aIndex = 0
            Select Case Mid(TELEBIB_CODE(X.Row - 1), 2, 1)
                Case "K"
                    SharedFl = TABLE_CUSTOMERS
                Case "L"
                    SharedFl = TABLE_SUPPLIERS
                Case "R"
                    SharedFl = TABLE_LEDGERACCOUNTS
                Case Else
                    MsgBox "nog niks"
            End Select
            GridText = ""
            If Mid(TELEBIB_CODE(X.Row - 1), 3, 2) <> "  " Then
                If DummyText <> "" Then
                    If SharedFl = TABLE_SUPPLIERS And Mid(TELEBIB_CODE(X.Row - 1), 3, 2) = "CO" Then
                        GridText = "CO" + DummyText
                    Else
                        GridText = DummyText
                    End If
                Else
                    GridText = Mid(TELEBIB_CODE(X.Row - 1), 3, 2) + "@Beperk@"
                End If
            Else
                GridText = DummyText
            End If
            SqlSearch.Show 1
            If Ktrl = 0 Then
                If SharedFl = TABLE_SUPPLIERS And Mid(TELEBIB_CODE(X.Row - 1), 3, 2) = "CO" Then
                    X.text = RTrim$(Mid(FVT(SharedFl, 0), 3))
                Else
                    X.text = FVT(SharedFl, 0)
                End If
            End If
        Case "  "
        Case Else
            Select Case Mid(TELEBIB_CODE(X.Row - 1), 1, 1)
                Case " "
                    BoxType = 0
                Case "0" To "9"
                    BoxType = 1
            End Select
            Select Case KeyCode
                Case 17
                    aIndex = Val(Mid(TELEBIB_CODE(X.Row - 1), 1, 3))
                    If BoxType = 1 Then
                        aIndex = aIndex + 1000
                    End If
                    X.Col = 2
                    DummyText = X.text
                    GridText = DummyText
                    KeuzeVSF.Show 1
                    If GridText <> DummyText Then
                        DummyText = GridText
                        X.text = DummyText
                    End If
            End Select
    End Select
End If

End Sub

Private Sub X_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
    Case 33 To 126
        If X.Col = 2 And Mid(X.Clip, 2, 2) = "  " Then
            X.text = Chr(KeyAscii) + X.text
            WijzigenLijn_Click
        End If
End Select

End Sub

Private Sub X_KeyUp(KeyCode As Integer, Shift As Integer)

X_Click

End Sub


Private Sub X_RowColChange()
    
    Dim ktrlHier As Boolean
    
    SSTab1.TabCaption(1) = "- Geen Bijlage"
    If SharedScanFl = 0 Or Trim$(X.TextMatrix(X.Row, 0)) = "" Then
    Else
        'bstAfbeeldingHier = App.path & "\" & X.TextMatrix(X.Row, 0) & ".tif"
        'bstAfbeeldingPDF = App.path & "\" & X.TextMatrix(X.Row, 0) & ".pdf"
        'Debug.Print SharedScanFl, X.Row, X.Col, X.TextMatrix(X.Row, 0)
        ktrlHier = ADO_GET(SharedScanFl, 0, "=", X.TextMatrix(X.Row, 0))
        If ktrlHier = True Then
            Err = 0
            Me.Label1.Caption = X.TextMatrix(X.Row, 0)
            On Local Error Resume Next
            If IsNull(rsMAR(SharedScanFl)("bstBLOB37")) Then
                Me.Label1.Caption = "- leeg -"
            Else
                SSTab1.TabCaption(1) = "+ Met Bijlage"
            End If
            If Err Then MsgBox Error
        End If
         
    End If
    
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
    DoEvents
    
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


