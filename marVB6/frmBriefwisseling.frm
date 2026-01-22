VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form VrijBericht 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vrije Berichtgeving"
   ClientHeight    =   5325
   ClientLeft      =   -4230
   ClientTop       =   2910
   ClientWidth     =   8700
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
   ScaleHeight     =   9.393
   ScaleMode       =   7  'Centimeter
   ScaleWidth      =   15.346
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
      Left            =   480
      Style           =   2  'Dropdown List
      TabIndex        =   41
      ToolTipText     =   "Dokumenten printer"
      Top             =   4920
      Width           =   6195
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
      Left            =   120
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4920
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
      Height          =   285
      Index           =   14
      Left            =   840
      TabIndex        =   39
      Text            =   "v"
      Top             =   1560
      Width           =   2625
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
      Left            =   840
      TabIndex        =   38
      Text            =   "v"
      Top             =   1200
      Width           =   2625
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
      Left            =   120
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1200
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
      Left            =   120
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1560
      Width           =   615
   End
   Begin RichTextLib.RichTextBox txtTekst0 
      Height          =   2475
      Left            =   60
      TabIndex        =   3
      Top             =   2340
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   4366
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmBriefwisseling.frx":030A
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
   Begin MSMAPI.MAPIMessages MPIBericht 
      Left            =   6900
      Top             =   2460
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
      Left            =   7620
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
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
      Top             =   1440
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
      Left            =   7440
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   4920
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
      Left            =   7560
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1800
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
      Left            =   6420
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
      Left            =   6420
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
      ItemData        =   "frmBriefwisseling.frx":0393
      Left            =   60
      List            =   "frmBriefwisseling.frx":0395
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   60
      Width           =   3465
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Dokumenttype"
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
      Left            =   3600
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
      Left            =   4350
      TabIndex        =   9
      Top             =   1080
      Width           =   2055
   End
   Begin VB.TextBox TxtTekst 
      Enabled         =   0   'False
      Height          =   285
      Index           =   4
      Left            =   3600
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
      Left            =   4350
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
      Left            =   3600
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
      Left            =   6510
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
      Left            =   7080
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
      Left            =   3600
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
      Left            =   4020
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
      Left            =   4950
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

Dim DDEFlag             As Integer
Dim DDE_vBib(13) As String
Dim Aantal_DDE          As Integer
Dim Tel                 As Integer

Dim WProgramma          As String
Dim PNaam               As String
Dim Wdokument           As String
Dim PrintSendKeys       As String
Dim CloseSendKeys       As String
Dim ReedsBestaand As Integer
Dim TekstIsRich As Boolean

Dim rsBrief As ADODB.Recordset
Dim wApp As Word.Application

Public g_objFaxDocument As New FAXCOMEXLib.FaxDocument
Public objFaxVerzender As FaxSender

Private Sub BerichtInOfUit_Click(Index As Integer)
Dim XXX As String


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
        Xlog.CmdAfdrukken.Visible = False
        Xlog.X.Row = 1
        Xlog.X.Col = 0

        Xlog.X.ColWidth(0) = 1680
        Xlog.X.ColWidth(1) = 4785
        Xlog.X.ColAlignment(0) = flexAlignLeftTop
        Xlog.X.ColAlignment(1) = flexAlignLeftTop
        XLogKey = ""
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
        End If
        rsBrief.Update
        Schoon
        Zoeken.SetFocus
        
End Select
Exit Sub

VolgendeLijn:
T = T + 1
Select Case FlPartij
    Case FlKlant, FlLeverancier
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


Private Function cbwAPP(printIT As Boolean, saveIT As Boolean) As String

    On Error GoTo wordAppErr
       
    Dim TT As Integer
    Dim volNaam As String
    
    cbwAPP = ""
    Set wApp = New Word.Application
    volNaam = BedrijfsLokatie & "sjb\sBrief.doc"
    If Dir(volNaam) = "" Then
        Set wApp = Nothing
        Exit Function
    Else
        wApp.Documents.Open (volNaam)
    End If
        
    Dim Book() As Variant
    ReDim Book(14, 2)
    Book(1, 1) = TxtTekst(2) '"Dhr."
    Book(2, 1) = TxtTekst(3) '"Van Den Plas"
    Book(3, 1) = TxtTekst(0) '"Armand"
    Book(4, 1) = TxtTekst(4) '"Mw."
    Book(5, 1) = TxtTekst(5) '"Regenbui"
    Book(6, 1) = TxtTekst(12) '"Inge"
    Book(7, 1) = TxtTekst(6) '"Populierenstraat"
    Book(8, 1) = TxtTekst(7) '39
    If TxtTekst(8) = "" Then
    Else
        Book(9, 1) = "b: " & TxtTekst(8) '"bus 8"
    End If
    Book(10, 1) = TxtTekst(9) '"B"
    Book(11, 1) = TxtTekst(10) '9310
    Book(12, 1) = TxtTekst(11) '"Moorsel"
    Book(13, 1) = TxtTekst(1) 'betreft
    Book(14, 1) = txtTekst0.Text
    'Book(14, 1) = txtTekst0.TextRTF
    
    
    Book(1, 2) = "Aan1"
    Book(2, 2) = "Naam1"
    Book(3, 2) = "Voornaam1"
    Book(4, 2) = "Aan2"
    Book(5, 2) = "Naam2"
    Book(6, 2) = "Voornaam2"
    Book(7, 2) = "Straat"
    Book(8, 2) = "Nr"
    Book(9, 2) = "Bus"
    Book(10, 2) = "LND"
    Book(11, 2) = "PC"
    Book(12, 2) = "Plaats"
    Book(13, 2) = "Betreft"
    Book(14, 2) = "TekstTest"
    
    With wApp.Documents(1)
        For TT = 1 To 14
            .Bookmarks(Book(TT, 2)).Select
            .Bookmarks(Book(TT, 2)).Range.Text = Book(TT, 1) ' .InsertAfter (Book(TT, 1))
        Next
    End With
    volNaam = BedrijfsLokatie & "sjb\sBrief" & Format(Now, "yyyymmddhhmmss") & ".doc"
    SnelHelpPrint wApp.Documents(1).Name
    If saveIT Then wApp.Documents(1).SaveAs volNaam
    If printIT Then wApp.Documents(1).PrintOut
    
    wApp.Documents(1).Close SaveChanges:=False
    cbwAPP = volNaam
    Set wApp = Nothing
    Exit Function

wordAppErr:
MsgBox Error
        
End Function

Private Sub cbDocumentPrinterHier_Click()

    dokumentPrinterNr = cbDocumentPrinterHier.ListIndex
    
End Sub

Private Sub cbTogglePrinter_Click()

     Me.cbDocumentPrinterHier.Enabled = Not Me.cbDocumentPrinterHier.Enabled
     
End Sub

Private Sub CmbAfdrukModus_Click()

Select Case Left(CmbAfdrukModus.Text, 1)
    Case "0", "1", "3"
        If MPISessie.SessionID <> 0 Then
        Else
            On Local Error Resume Next
            MPISessie.SignOn
            If Err Then
                'Msg = "U kiest om het mailsysteem niet aan te wenden." & vbCr
                'Msg = Msg & "Probeer eventueel later opnieuw E-mail aan te duiden" & vbCr & vbCr
                'Msg = Msg & "Foutmelding van het systeem: " & Err.Description
                'MsgBox Msg, vbExclamation
                SnelHelpPrint Err.Description
                CmbAfdrukModus.ListIndex = 2
            Else
                On Local Error GoTo 0
                MPIBericht.SessionID = MPISessie.SessionID
                SnelHelpPrint "E-mail sessie met succes opgestart. IDkode :" + Format(MPISessie.SessionID)
            End If
        End If
    Case "2"
        If MPISessie.SessionID = 0 Then
        Else
            On Local Error Resume Next
            MPISessie.SignOff
            If Err Then
                MsgBox "Fout tijdens afsluiten van een sessie...  Postvak In (InBox) werd voortijdig gesloten."
            Else
                SnelHelpPrint "E-mail sessie werd met succes beëindigd..."
            End If
        End If
    Case Else
        MsgBox "stop"
End Select

End Sub


Private Sub cmdEmailen_Click()

Dim BestandHier As String
If MPISessie.SessionID <> 0 Then
    On Local Error GoTo MPIError
    'Compose new message
    MPIBericht.Compose
    
    'Address message
    If CmbAfdrukModus.ListIndex = 0 Then
        MPIBericht.RecipDisplayName = "Roelandt&Vermoesen"
        MPIBericht.RecipAddress = "SMTP:jos@rv.be"
    Else
        MPIBericht.RecipDisplayName = rsMAR(FlPartij)("A100")
        MPIBericht.RecipAddress = "SMTP:" & rsMAR(FlPartij)("v224")
    End If
    'Resolve recipient name
    MPIBericht.AddressResolveUI = True '=dialogbox, false = error genereren
    'MpiBericht.ResolveName
        
    'Create the message
    If CmbAfdrukModus.ListIndex = 0 Then
        MPIBericht.MsgSubject = "$$/" & usrMailAdres & "/$$"
    Else
        MPIBericht.MsgSubject = TxtTekst(1).Text
    End If
    MPIBericht.MsgNoteText = "Briefwisseling in bijlage als MS Word Document" 'txtTekst0.Text
    KtrlBox = MsgBox("Verstuurde E-mail aan " & vBibTekst(FlPartij, "#v224 #") & vbCr & vbCr & "Afdruk maken ?", vbQuestion + vbYesNo + vbDefaultButton2)
    If KtrlBox = vbYes Then
        BestandHier = cbwAPP(True, True)
    Else
        BestandHier = cbwAPP(False, True)
    End If
    MPIBericht.AttachmentPathName = BestandHier
        
    'Send the message
    MPIBericht.Send True 'False
    BerichtInOfUit_Click 1
    Exit Sub
Else
    MsgBox "Uw E-mail systeem is niet geactiveerd.  Contacteer uw systeembeheerder voor bijkomende inlichtingen.", vbExclamation
End If
Exit Sub

MPIError:
MsgBox Error
Resume Next

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
     Set objFaxVerzender = g_objFaxDocument.Sender
     If Err Then
        MsgBox "Enkel beschikbaar voor Windows XP en hoger", vbInformation
        Screen.MousePointer = vbNormal
        Exit Sub
     End If
     objFaxVerzender.LoadDefaultSender
     
'Error handling
    'On Error GoTo Error_Handler
    
    'Check if the Aanhangsel checkbox was selected, but no Aanhangsel provided
    'If chkAanhangsel.Value = 1 And txtAanhangselPath.Text = "" Then
    '    MsgBox "Duidt het path voor de bijlage aan en probeer opnieuw.", vbExclamation
    '    Exit Sub
    'End If

    'Remove the previous recipient, if there is one. Otherwise, the fax would be sent to
    'recipients that accumulate in the FaxRecipients collection from each previous use of the form.
    If g_objFaxDocument.Recipients.Count = 1 Then g_objFaxDocument.Recipients.Remove (1)
    
    'Set no cover page as default
    g_objFaxDocument.CoverPageType = fcptNONE
        
    'If the checkbox is selected, set cover page type to "local"
    'and set the path and file name to the textbox text
    'If chkVoorBlad.Value = 1 Then
    '    g_objFaxDocument.CoverPageType = fcptLOCAL
    '    g_objFaxDocument.CoverPage = txtCPPath.Text
    'End If
    
    'Set the rest of the fax document properties
    'Noties voorblad
    g_objFaxDocument.Note = "" 'txtNotities.Text
    g_objFaxDocument.Subject = "" 'txtOnderwerp.Text
    g_objFaxDocument.Recipients.Add TxtTekst(13).Text, TxtTekst(3).Text   'txtFaxNummerGeadresseerde.Text, txtNaamGeadresseerde.Text
    g_objFaxDocument.Sender.Name = objFaxVerzender.Name
    g_objFaxDocument.Sender.FaxNumber = objFaxVerzender.FaxNumber
    g_objFaxDocument.Sender.OfficePhone = objFaxVerzender.OfficePhone
    'Include the Aanhangsel if the Aanhangsel checkbox is selected
    'If chkAanhangsel.Value = 1 Then g_objFaxDocument.Body = txtAanhangselPath.Text
        
    Dim BestandHier As String
    If cbLokaleAfdruk.Value = 0 Then
        BestandHier = cbwAPP(False, True)
    Else
        BestandHier = cbwAPP(True, True)
    End If
    g_objFaxDocument.Body = BestandHier 'txtAanhangselPath.Text
    'g_objFaxDocument.Body = "D:\Mijn Documenten\MODI\test.xls"
     
    'Submit the document to the server
    Call g_objFaxDocument.Submit("") '(ServerNameText.Text)
    Kill BestandHier
    If TxtTekst(13).Text <> txtFaxIs Then
        MsgBox "gewijzigd nummer.  Hier nog routine vraagstelling om al dan niet de partijFiche bij te werken", vbInformation
    End If
    BerichtInOfUit_Click 1
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub cmdMail_Click()

    If cmdMail.Enabled = False Then
        Exit Sub
    Else
        Screen.MousePointer = vbHourglass
    End If
    
    Dim BestandHier As String
    If MPISessie.SessionID <> 0 Then
    On Local Error GoTo MPIError
    'Compose new message
    MPIBericht.Compose
    
    'Address message
    If CmbAfdrukModus.ListIndex = 0 Then
        MPIBericht.RecipDisplayName = rsMAR(FlPartij)("A100")
        MPIBericht.RecipAddress = "SMTP:" & TxtTekst(14).Text
    Else
        MPIBericht.RecipDisplayName = rsMAR(FlPartij)("A100")
        MPIBericht.RecipAddress = "SMTP:" & TxtTekst(14).Text
    End If
    'Resolve recipient name
    MPIBericht.AddressResolveUI = True '=dialogbox, false = error genereren
    'MpiBericht.ResolveName
        
    'Create the message
    MPIBericht.MsgSubject = TxtTekst(1).Text
    If CmbAfdrukModus.ListIndex = 0 Then
        MPIBericht.MsgNoteText = txtTekst0.Text
    Else
        MPIBericht.MsgSubject = TxtTekst(1).Text
        MPIBericht.MsgNoteText = "Briefwisseling in bijlage als MS Word Document" 'txtTekst0.Text
        'KtrlBox = MsgBox("Verstuurde E-mail aan " & vBibTekst(FlPartij, "#v224 #") & vbCr & vbCr & "Afdruk maken ?", vbQuestion + vbYesNo + vbDefaultButton2)
        If cbLokaleAfdruk.Value = 1 Then
            BestandHier = cbwAPP(True, True)
        Else
            BestandHier = cbwAPP(False, True)
        End If
        MPIBericht.AttachmentPathName = BestandHier
    End If
    'Send the message
    On Error Resume Next
    MPIBericht.Send False
    If Err = 32001 Then Screen.MousePointer = vbNormal: Exit Sub
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

Private Sub cmdPost_Click()

    Stop
    
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
For TelTot = 0 To Printers.Count - 1
    Me.cbDocumentPrinterHier.AddItem Printers(TelTot).Port + " " + Printers(TelTot).DeviceName
Next
cbDocumentPrinterHier.ListIndex = dokumentPrinterNr

PaginaTeller = 0
Partij(0).Value = True
txtTekst0.Font = "Courier New"
MaxLijn79 = 90
Briefwisseling(1).Enabled = True
Briefwisseling(0).Value = True

KontaktPersoon = Val(String99(Lees, 201))
DDEFlag = Val(String99(Lees, 290))
If DDEFlag = 1 Then
    OpenDDE
End If

CmbAfdrukModus.Clear
CmbAfdrukModus.AddItem "0: E-mail via normale tekst"
CmbAfdrukModus.AddItem "1: E-mail via insluiting van Word Sjabloon"
CmbAfdrukModus.AddItem "2: E-mail wordt niet gebruikt op dit ogenblik"
If Left(String99(Lees, 21), 1) = "1" Then
    CmbAfdrukModus.ListIndex = 0
Else
    CmbAfdrukModus.ListIndex = 2
    CmbAfdrukModus.Enabled = False
End If
Schoon

End Sub

Private Sub Form_Unload(Cancel As Integer)

If MPISessie.SessionID = 0 Then
Else
    On Local Error Resume Next
    MPISessie.SignOff
    If Err Then
        MsgBox "Fout tijdens afsluiten van een sessie...  Postvak In (InBox) werd voortijdig gesloten."
    End If
End If

If DDEFlag = 1 Then
    SluitDDE
End If

End Sub

Private Sub OpenDDE()

On Local Error Resume Next

WProgramma = String99(Lees, 291)
PNaam = String99(Lees, 292)
Wdokument = String99(Lees, 293)
PrintSendKeys = String99(Lees, 294)
CloseSendKeys = String99(Lees, 295)
Caption = Caption + " via [" + PNaam + " - " + Wdokument + "]"
'This procedure starts WINWORD.EXE, loads the dokument that was
'created earlier, and prepared for DDE by creating a bookmark to
'the whole dokument. This bookmark is necessary because it
'functions as the LinkItem for the source in the DDE conversation.
'z% = Shell("WinWord c:\marnt\ntbrief.Doc", 1)
Err = 0
Z% = Shell(WProgramma + " " + Wdokument, 1)
If Err Then
    MsgBox Error$ + vbCrLf + vbCrLf + "Vermoedelijk rond opdracht : " + WProgramma + " " + Wdokument
    Exit Sub
End If
XDoEvents = DoEvents()
'Process Windows events to ensure that
'WINWORD.EXE is executed before any attempt is
'made to perform DDE with it.

DDE_vBib(0) = "v132"
DDE_vBib(1) = "v129"
DDE_vBib(2) = "A102"
DDE_vBib(3) = "A100"
DDE_vBib(4) = "vs01"
DDE_vBib(5) = "A125"
DDE_vBib(6) = "A104"
DDE_vBib(7) = "A105"
DDE_vBib(8) = "A106"
DDE_vBib(9) = "A109"
DDE_vBib(10) = "A107"
DDE_vBib(11) = "A108"

Aantal_DDE = 11
For Tel = 0 To Aantal_DDE
    TxtTekst(Tel).LinkMode = 0                 'Clears DDE link if it exists.
    TxtTekst(Tel).LinkTopic = WProgramma + "|" + Wdokument 'Sets up link with WINWORD.EXE.
    TxtTekst(Tel).LinkItem = "DDE_" + DDE_vBib(Tel)   'Set link to bookmark on dokument.
    TxtTekst(Tel).LinkMode = 2                 'Establish a manual DDE link.
Next

End Sub


Private Sub Partij_click(Index As Integer)

Select Case Partij(0).Value
    Case True
        PartijKode = "K"
        FlPartij = FlKlant
    Case False
        PartijKode = "L"
        FlPartij = FlLeverancier
End Select
BriefType = "B"
Briefwisseling(0).Value = 1

Sortering.Clear
For T = 0 To FlAantalIndexen(FlPartij)
    Sortering.AddItem Format$(T, "00") + ":" + FLIndexCaption(FlPartij, T)
Next
Sortering.ListIndex = 1

End Sub

Private Sub PokeIt()

On Local Error GoTo ErIsEenFout
'With any DDE link, this button is visible. Clicking this button
'pokes information from the destination application into the source
'application.
For Tel = 0 To Aantal_DDE
    TxtTekst(Tel).LinkPoke
Next
AppActivate PNaam
SendKeys PrintSendKeys, True
XDoEvents = DoEvents
VrijBericht.SetFocus

Exit Sub

ErIsEenFout:
'MsgBox Error + vbCrLf + vbCrLf + "Onverwachte situatie tijdens de DDE-conversatie.  Kontroleer uw programma's en/of dokumentsjablonen !"
Exit Sub
Resume

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

If Dir$(BedrijfsLokatie + "DDEF" + TypeEnTaal + ".Txt") = "" Then
    Beep
    Exit Sub
Else
    FlFree = FreeFile
    Open BedrijfsLokatie + "DDEF" + TypeEnTaal + ".Txt" For Input As FlFree
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

If Dir$(BedrijfsLokatie + "DDEF" + TypeEnTaal + "G.Txt") = "" Then
Else
    FlFree = FreeFile
    Open BedrijfsLokatie + "DDEF" + TypeEnTaal + "G.Txt" For Input As FlFree
    While Not EOF(FlFree)
        Input #FlFree, FiguurX, FiguurY, FiguurName
        If ScrLeesTekstBestand(Msg, BedrijfsLokatie & FiguurName & ".mfd") Then
            FigBestandsnaam = Left(Msg, InStr(Msg, vbTab) - 1)
            Mim.imgFiguur.Picture = LoadPicture(FigBestandsnaam)
            Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
            Hoogte = Left(Msg, InStr(Msg, vbTab) - 1)
            Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
            Breedte = Msg
            Printer.PaintPicture Mim.imgFiguur.Picture, FiguurX, FiguurY, Breedte, Hoogte
        Else
            MsgBox "probleemke"
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

Wegschrijven.Enabled = False
txtTekst0.Enabled = False
BerichtInofUit(1).Enabled = False
BerichtInofUit(0).Enabled = True
Zoeken.Enabled = True
TLBRecord(FlAllerlei) = ""

Partij_click 0

End Sub

Private Sub SchoonVegen_Click()

Schoon

End Sub

Private Sub SluitDDE()

On Local Error GoTo NietMeer

If DDEFlag = 1 Then
    AppActivate PNaam
    SendKeys CloseSendKeys, True
End If
Exit Sub

NietMeer:
MsgBox "De andere toepassing is reeds afgesloten !"
Exit Sub
Resume

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
        Else
            cmdMail.Enabled = True
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
ElseIf DDEFlag <> 1 Then
    TelLijn = 0
    Tekst = txtTekst0.Text
    Do While Tekst <> ""
        positieCRLF = InStr(Tekst, vbCrLf)
        If positieCRLF Then
            TekstLijn = Left$(Tekst, positieCRLF - 1)
            Tekst = Right$(Tekst, Len(Tekst) - positieCRLF - 1)
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
DrukAf

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
If Trim(rsMAR(FlPartij)("v224")) <> "" Then   'email
    txtMailIs = rsMAR(FlPartij)("v224")
    TxtTekst(14).Text = txtMailIs
Else
    txtMailIs = ""
    TxtTekst(14).Text = ""
End If
If Trim(rsMAR(FlPartij)("vs02")) <> "" Then   'fax
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
    TxtTekst(2) = Mid$(fmarBoxText("003", Taal, vBibTekst(FlPartij, "#A102 #")), 4, 10)
    TxtTekst(3) = vBibTekst(FlPartij, "#A100 #")
    TxtTekst(0) = vBibTekst(FlPartij, "#A101 #")
    If Val(vBibTekst(FlPartij, "#A102 #")) = 0 Then
        TxtTekst(2) = ""
    End If
    TxtTekst(4) = Mid$(fmarBoxText("003", Taal, vBibTekst(FlPartij, "#vs01 #")), 4, 10)
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
            MsgBox "Partij '" & Trim((VanTot(2).Caption)) & "' niet gevonden.  Gelieve klant of leveranciers nogmaals te SELECTeren en opnieuw proberen afdrukken a.u.b."
        Else
            RecordToVeld FlPartij
            GoSub PrintHetdokument
        End If
    Case False
        bGetOrGreater FlPartij, (Sortering.ListIndex), vSet((VanTot(0).Caption), FlIndexLen(FlPartij, (Sortering.ListIndex)))
        If Ktrl Or KeyBuf(FlPartij) > vSet((VanTot(1).Caption), FlIndexLen(FlPartij, (Sortering.ListIndex))) Then
        Else
            RecordToVeld FlPartij
            GoSub PrintHetdokument
            Do
                bNext FlPartij
                If Ktrl Or KeyBuf(FlPartij) > vSet((VanTot(1).Caption), FlIndexLen(FlPartij, (Sortering.ListIndex))) Then
                    Exit Do
                Else
                    RecordToVeld FlPartij
                    GoSub PrintHetdokument
                End If
            Loop
        End If
    Case Else
        MsgBox "Partij '" & Trim((VanTot(2).Caption)) & "' niet gevonden.  Gelieve klant of leveranciers nogmaals te SELECTeren en opnieuw proberen afdrukken a.u.b."
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
TxtTekst(2) = Mid$(fmarBoxText("003", Taal, vBibTekst(FlPartij, "#A102 #")), 4, 10)
TxtTekst(3) = vBibTekst(FlPartij, "#A100 #")
TxtTekst(0) = vBibTekst(FlPartij, "#A101 #")
If Val(vBibTekst(FlPartij, "#A102 #")) = 0 Then
    rSip(0) = TxtTekst(3)
    TxtTekst(2) = ""
Else
    rSip(0) = Trim(TxtTekst(2)) + " " _
            + Trim(TxtTekst(3)) + " " _
            + Trim(TxtTekst(0))
End If
TxtTekst(4) = Mid(fmarBoxText("003", Taal, vBibTekst(FlPartij, "#vs01 #")), 4, 10)
TxtTekst(5) = vBibTekst(FlPartij, "#A125 #")
TxtTekst(12) = vBibTekst(FlPartij, "#A127 #")
If KontaktPersoon = 1 Then
    If Val(vBibTekst(FlPartij, "#vs01 #")) = 0 Then
        rSip(1) = TxtTekst(5)
        TxtTekst(4) = ""
    Else
        rSip(1) = Trim(TxtTekst(4)) + " " _
                + Trim(TxtTekst(5)) + " " _
                + Trim(TxtTekst(12))
    End If
Else
    rSip(1) = ""
    TxtTekst(4) = ""
    TxtTekst(5) = ""
End If
TxtTekst(6) = vBibTekst(FlPartij, "#A104 #")
TxtTekst(7) = vBibTekst(FlPartij, "#A105 #")
TxtTekst(8) = vBibTekst(FlPartij, "#A106 #")
rSip(2) = Trim(TxtTekst(6)) + " " _
        + Trim(TxtTekst(7)) + " " _
        + Trim(TxtTekst(8))

TxtTekst(9) = vBibTekst(FlPartij, "#A109 #")
TxtTekst(10) = vBibTekst(FlPartij, "#A107 #")
TxtTekst(11) = vBibTekst(FlPartij, "#A108 #")
rSip(4) = TxtTekst(9) + " " _
        + TxtTekst(10) + " " _
        + TxtTekst(11)

If DDEFlag = 1 Then
    PokeIt
    XDoEvents = DoEvents
    Return
End If
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
If Printer.DeviceName = "Fax" Then
    Printer.FontBold = True
Else
    Printer.FontBold = False
End If
Return

DetailLijnen:
Do While Tekst <> ""
    positieCRLF = InStr(Tekst, vbCrLf)
    If positieCRLF Then
        Printer.Print Tab(TabIs); Left$(Tekst, positieCRLF - 1); vbCrLf;
        Tekst = Right$(Tekst, Len(Tekst) - positieCRLF - 1)
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
MsgBox "Onverwachte fout bij openen briefhoofdbestand." + vbCrLf + vbCrLf + "Stopkode : " + Format$(Err) + vbCrLf + vbCrLf + "Rapporteer aan R&&Vsoft en/of kontroleer uw setup."
Ktrl = 100
Exit Sub

End Sub


