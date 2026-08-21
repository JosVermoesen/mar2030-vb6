VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGroeneKaart 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Groene Kaart Auto"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   Icon            =   "frmGroeneKaart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   9075
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdAfdrukken 
      Caption         =   "Afdruk op &Sjabloon  van verzekeraar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7560
      TabIndex        =   73
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CheckBox cbKleur 
      Caption         =   "Kleur"
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   5160
      Width           =   735
   End
   Begin VB.ComboBox cbKwaliteit 
      Height          =   315
      ItemData        =   "frmGroeneKaart.frx":0442
      Left            =   1080
      List            =   "frmGroeneKaart.frx":0456
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cbTogglePrinter 
      Caption         =   "..."
      Height          =   315
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4320
      Width           =   315
   End
   Begin VB.ComboBox cbDocumentPrinterHier 
      Enabled         =   0   'False
      Height          =   315
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   4
      ToolTipText     =   "Dokumenten printer"
      Top             =   4320
      Width           =   4635
   End
   Begin VB.CommandButton cmdAfdruk 
      Caption         =   "Grafische &Afdruk op blanco groen papier.."
      Height          =   1005
      Left            =   5520
      Picture         =   "frmGroeneKaart.frx":047D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5280
      Width           =   1965
   End
   Begin VB.CommandButton cbSchoon 
      Caption         =   "Schoon"
      Height          =   300
      Left            =   6960
      TabIndex        =   2
      Top             =   240
      Width           =   885
   End
   Begin VB.CommandButton cmdHistoriek 
      Caption         =   "&Historiek"
      Height          =   300
      Left            =   5520
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   240
      Width           =   1240
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   300
      Left            =   8040
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   240
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3975
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   7011
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Manueel"
      TabPicture(0)   =   "frmGroeneKaart.frx":08BF
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblInfo(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblInfo(8)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblInfo(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblInfo(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblInfo(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblInfo(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblInfo(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblInfo(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblInfo(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblInfo(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dtpTot"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "mebTekstinfo(5)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "mebTekstinfo(4)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "mebTekstinfo(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "mebTekstinfo(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "mebTekstinfo(1)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "mebTekstinfo(0)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "dtpVan"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmbSoort"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmbMaatschappij"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "scanBeheer"
      TabPicture(1)   =   "frmGroeneKaart.frx":08DB
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdUpdate"
      Tab(1).Control(1)=   "cmdScan"
      Tab(1).Control(2)=   "cmdScanOpties"
      Tab(1).Control(3)=   "cmdScanNieuweAfbeelding"
      Tab(1).Control(4)=   "cmdBestand"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdBewaar"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "tbGK(2)"
      Tab(1).Control(7)=   "tbGK(1)"
      Tab(1).Control(8)=   "tbGK(0)"
      Tab(1).Control(9)=   "cmdTop"
      Tab(1).Control(10)=   "cmdVorige"
      Tab(1).Control(11)=   "cmdVolgende"
      Tab(1).Control(12)=   "cmdBodem"
      Tab(1).ControlCount=   13
      TabCaption(2)   =   "XY Instellingen"
      TabPicture(2)   =   "frmGroeneKaart.frx":08F7
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblInfo(10)"
      Tab(2).Control(1)=   "lblInfo(11)"
      Tab(2).Control(2)=   "lblInfo(12)"
      Tab(2).Control(3)=   "lblInfo(22)"
      Tab(2).Control(4)=   "lblInfo(21)"
      Tab(2).Control(5)=   "lblInfo(20)"
      Tab(2).Control(6)=   "lblInfo(19)"
      Tab(2).Control(7)=   "lblInfo(18)"
      Tab(2).Control(8)=   "lblInfo(17)"
      Tab(2).Control(9)=   "lblInfo(16)"
      Tab(2).Control(10)=   "lblInfo(15)"
      Tab(2).Control(11)=   "lblInfo(14)"
      Tab(2).Control(12)=   "lblInfo(13)"
      Tab(2).Control(13)=   "mebXYPositie(9)"
      Tab(2).Control(14)=   "mebXYPositie(5)"
      Tab(2).Control(15)=   "mebXYPositie(6)"
      Tab(2).Control(16)=   "mebXYPositie(4)"
      Tab(2).Control(17)=   "mebXYPositie(8)"
      Tab(2).Control(18)=   "mebXYPositie(7)"
      Tab(2).Control(19)=   "mebXYPositie(3)"
      Tab(2).Control(20)=   "mebXYPositie(2)"
      Tab(2).Control(21)=   "mebXYPositie(1)"
      Tab(2).Control(22)=   "mebXYPositie(0)"
      Tab(2).Control(23)=   "cbXYwijzigen"
      Tab(2).Control(24)=   "cmbFormaatVan"
      Tab(2).Control(25)=   "cmbFormaatTot"
      Tab(2).Control(26)=   "cmdWijzigen"
      Tab(2).Control(26).Enabled=   0   'False
      Tab(2).Control(27)=   "cmdBewaren"
      Tab(2).Control(27).Enabled=   0   'False
      Tab(2).ControlCount=   28
      TabCaption(3)   =   "Lease Connect Import"
      TabPicture(3)   =   "frmGroeneKaart.frx":0913
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "mfgLijst"
      Tab(3).Control(1)=   "cbLeaseConnect"
      Tab(3).ControlCount=   2
      Begin VB.ComboBox cmbMaatschappij 
         Height          =   315
         ItemData        =   "frmGroeneKaart.frx":092F
         Left            =   1320
         List            =   "frmGroeneKaart.frx":0939
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   480
         Width           =   3705
      End
      Begin VB.ComboBox cmbSoort 
         Height          =   315
         ItemData        =   "frmGroeneKaart.frx":097D
         Left            =   1320
         List            =   "frmGroeneKaart.frx":0993
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   2640
         Width           =   3705
      End
      Begin VB.CommandButton cmdBodem 
         Caption         =   "Bodem"
         Height          =   495
         Left            =   -70560
         TabIndex        =   28
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton cmdVolgende 
         Caption         =   ">"
         Height          =   495
         Left            =   -70200
         TabIndex        =   27
         Top             =   1980
         Width           =   255
      End
      Begin VB.CommandButton cmdVorige 
         Caption         =   "<"
         Height          =   495
         Left            =   -70560
         TabIndex        =   26
         Top             =   1980
         Width           =   255
      End
      Begin VB.CommandButton cmdTop 
         Caption         =   "Top"
         Height          =   495
         Left            =   -70560
         TabIndex        =   25
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox tbGK 
         Height          =   315
         Index           =   0
         Left            =   -74340
         TabIndex        =   24
         Tag             =   "tbSleutel"
         ToolTipText     =   "Code Verzekeraar"
         Top             =   1980
         Width           =   1095
      End
      Begin VB.TextBox tbGK 
         Height          =   315
         Index           =   1
         Left            =   -74340
         TabIndex        =   23
         Tag             =   "tbOmschrijving"
         ToolTipText     =   "Naam Verzekeraar"
         Top             =   2340
         Width           =   3675
      End
      Begin VB.TextBox tbGK 
         Height          =   315
         Index           =   2
         Left            =   -74340
         TabIndex        =   22
         Tag             =   "tbDocRef"
         ToolTipText     =   "Referte document Verzekeraar"
         Top             =   2700
         Width           =   3675
      End
      Begin VB.CommandButton cmdBewaar 
         Caption         =   "Bewaren"
         Height          =   375
         Left            =   -73800
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   420
         Width           =   975
      End
      Begin VB.CommandButton cmdBestand 
         Caption         =   "Bestand"
         Height          =   375
         Left            =   -74880
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   420
         Width           =   975
      End
      Begin VB.CommandButton cmdScanNieuweAfbeelding 
         Caption         =   "Afbeelding Scannen"
         Height          =   795
         Left            =   -71220
         TabIndex        =   19
         Top             =   420
         Width           =   1275
      End
      Begin VB.CommandButton cmdScanOpties 
         Caption         =   "Scan Opties"
         Height          =   375
         Left            =   -72600
         TabIndex        =   18
         Top             =   420
         Width           =   1335
      End
      Begin VB.CommandButton cmdScan 
         Caption         =   "Software Scan"
         Height          =   375
         Left            =   -72600
         TabIndex        =   17
         Top             =   840
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   315
         Left            =   -71760
         TabIndex        =   16
         Top             =   1980
         Width           =   1095
      End
      Begin VB.CommandButton cmdBewaren 
         Caption         =   "&Instellingen bewaren"
         Height          =   405
         Left            =   -72180
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   420
         Width           =   2205
      End
      Begin VB.CommandButton cmdWijzigen 
         Caption         =   "&Wijzigen"
         Height          =   405
         Left            =   -72180
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   900
         Width           =   2205
      End
      Begin VB.ComboBox cmbFormaatTot 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmGroeneKaart.frx":0A09
         Left            =   -71760
         List            =   "frmGroeneKaart.frx":0A19
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   3240
         Width           =   1785
      End
      Begin VB.ComboBox cmbFormaatVan 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmGroeneKaart.frx":0A53
         Left            =   -71760
         List            =   "frmGroeneKaart.frx":0A63
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2940
         Width           =   1785
      End
      Begin VB.CommandButton cbLeaseConnect 
         Caption         =   "Import"
         Height          =   375
         Left            =   -70800
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cbXYwijzigen 
         Caption         =   "Wijzigen via Kladblok"
         Height          =   495
         Left            =   -72120
         TabIndex        =   10
         Top             =   1440
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker dtpVan 
         Height          =   285
         Left            =   1320
         TabIndex        =   31
         Top             =   2970
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   503
         _Version        =   393216
         Format          =   73990144
         UpDown          =   -1  'True
         CurrentDate     =   43800
         MaxDate         =   47512
         MinDate         =   43466
      End
      Begin MSMask.MaskEdBox mebTekstinfo 
         Height          =   300
         Index           =   0
         Left            =   1320
         TabIndex        =   32
         Top             =   780
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebTekstinfo 
         Height          =   300
         Index           =   1
         Left            =   1320
         TabIndex        =   33
         Top             =   1080
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebTekstinfo 
         Height          =   300
         Index           =   2
         Left            =   1320
         TabIndex        =   34
         Top             =   1380
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebTekstinfo 
         Height          =   300
         Index           =   3
         Left            =   1320
         TabIndex        =   35
         Top             =   1680
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebTekstinfo 
         Height          =   300
         Index           =   4
         Left            =   1320
         TabIndex        =   36
         Top             =   1980
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebTekstinfo 
         Height          =   300
         Index           =   5
         Left            =   1320
         TabIndex        =   37
         Top             =   2280
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   529
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpTot 
         Height          =   315
         Left            =   1320
         TabIndex        =   38
         Top             =   3270
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   556
         _Version        =   393216
         Format          =   73990144
         UpDown          =   -1  'True
         CurrentDate     =   45444
         MaxDate         =   47848
         MinDate         =   45444
      End
      Begin MSMask.MaskEdBox mebXYPositie 
         Height          =   285
         Index           =   0
         Left            =   -73500
         TabIndex        =   39
         Top             =   720
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   7
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebXYPositie 
         Height          =   285
         Index           =   1
         Left            =   -73500
         TabIndex        =   40
         Top             =   1020
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   7
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebXYPositie 
         Height          =   285
         Index           =   2
         Left            =   -73500
         TabIndex        =   41
         Top             =   1320
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   7
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebXYPositie 
         Height          =   285
         Index           =   3
         Left            =   -73500
         TabIndex        =   42
         Top             =   1620
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   7
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebXYPositie 
         Height          =   285
         Index           =   7
         Left            =   -73500
         TabIndex        =   43
         Top             =   2910
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   7
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebXYPositie 
         Height          =   285
         Index           =   8
         Left            =   -73500
         TabIndex        =   44
         Top             =   3240
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   7
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebXYPositie 
         Height          =   285
         Index           =   4
         Left            =   -73500
         TabIndex        =   45
         Top             =   1920
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   7
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebXYPositie 
         Height          =   285
         Index           =   6
         Left            =   -73500
         TabIndex        =   46
         Top             =   2580
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   7
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebXYPositie 
         Height          =   285
         Index           =   5
         Left            =   -73500
         TabIndex        =   47
         Top             =   2220
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   7
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebXYPositie 
         Height          =   285
         Index           =   9
         Left            =   -70920
         TabIndex        =   48
         Top             =   2340
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   503
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   16777215
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   7
         Mask            =   "###/###"
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgLijst 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   49
         Top             =   960
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   5106
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColor       =   0
         FixedCols       =   0
         FocusRect       =   2
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Naam Klant"
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   72
         Top             =   780
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Straat/Nummer"
         Height          =   285
         Index           =   1
         Left            =   60
         TabIndex        =   71
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Woonplaats"
         Height          =   285
         Index           =   2
         Left            =   60
         TabIndex        =   70
         Top             =   1380
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Polisnummer"
         Height          =   285
         Index           =   3
         Left            =   60
         TabIndex        =   69
         Top             =   1680
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maatschappij"
         Height          =   285
         Index           =   4
         Left            =   60
         TabIndex        =   68
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Geldig VAN"
         Height          =   285
         Index           =   5
         Left            =   60
         TabIndex        =   67
         Top             =   2970
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Geldig TOT"
         Height          =   315
         Index           =   6
         Left            =   60
         TabIndex        =   66
         Top             =   3270
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kenteken"
         Height          =   285
         Index           =   7
         Left            =   60
         TabIndex        =   65
         Top             =   1980
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Soort"
         Height          =   315
         Index           =   8
         Left            =   60
         TabIndex        =   64
         Top             =   2640
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Merk"
         Height          =   285
         Index           =   9
         Left            =   60
         TabIndex        =   63
         Top             =   2280
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Merk"
         Height          =   285
         Index           =   13
         Left            =   -74880
         TabIndex        =   62
         Top             =   2220
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Soort"
         Height          =   315
         Index           =   14
         Left            =   -74880
         TabIndex        =   61
         Top             =   2580
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Kenteken"
         Height          =   285
         Index           =   15
         Left            =   -74880
         TabIndex        =   60
         Top             =   1920
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Geldig TOT"
         Height          =   315
         Index           =   16
         Left            =   -74880
         TabIndex        =   59
         Top             =   3210
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Geldig VAN"
         Height          =   285
         Index           =   17
         Left            =   -74880
         TabIndex        =   58
         Top             =   2910
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Maatschappij"
         Height          =   285
         Index           =   18
         Left            =   -74880
         TabIndex        =   57
         Top             =   420
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Polisnummer"
         Height          =   285
         Index           =   19
         Left            =   -74880
         TabIndex        =   56
         Top             =   1620
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Woonplaats"
         Height          =   285
         Index           =   20
         Left            =   -74880
         TabIndex        =   55
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Straat/Nummer"
         Height          =   285
         Index           =   21
         Left            =   -74880
         TabIndex        =   54
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Naam Klant"
         Height          =   285
         Index           =   22
         Left            =   -74880
         TabIndex        =   53
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "xy-Pos (mm.)"
         Height          =   285
         Index           =   12
         Left            =   -73500
         TabIndex        =   52
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Datummaskers Van/Tot"
         Height          =   285
         Index           =   11
         Left            =   -71760
         TabIndex        =   51
         Top             =   2640
         Width           =   1785
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "XYverlies"
         Height          =   285
         Index           =   10
         Left            =   -71760
         TabIndex        =   50
         Top             =   2340
         Width           =   825
      End
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Afdruk Kwaliteit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   120
      TabIndex        =   6
      Top             =   4800
      Width           =   915
   End
   Begin VB.Image imgFiguur 
      BorderStyle     =   1  'Fixed Single
      Height          =   4575
      Left            =   5520
      Stretch         =   -1  'True
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "frmGroeneKaart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cnn As ADODB.Connection
Dim rs  As ADODB.Recordset
Dim BestandsNaamFiguur As String
Dim GKIsKLAAR As Boolean


Private Sub cbDocumentPrinterHier_Click()

    dokumentPrinterNr = cbDocumentPrinterHier.ListIndex
        
End Sub

Private Sub cbLeaseConnect_Click()

    Dim strTB2() As String
    Dim strData() As String
    Dim flHier As Integer
    Dim COUNT_TO As Integer
    Dim strEenString As String
            
    On Error GoTo CancelError
    Mim.Teken.InitDir = App.path
    Mim.Teken.fileName = ""
    Mim.Teken.CancelError = True
    Mim.Teken.Filter = _
        "Lease Connect bestanden (IMAT*.csv)|IMAT*.csv"
        
    Mim.Teken.ShowOpen
    flHier = FreeFile
    Err = 0
    On Error Resume Next
    Open Mim.Teken.fileName For Input As flHier
        Line Input #flHier, strEenString
        strTB2 = Split(strEenString, ";")
        Line Input #flHier, strEenString
        strData = Split(strEenString, ";")
    Close flHier
    If Err Then MsgBox Error, vbExclamation, "Foutopvang bestandsysteem": Exit Sub
    If UBound(strTB2) <> 42 Then MsgBox "Andere versie Layout retourblok.  Installeer recente versie van GroeneKaart a.u.b.", vbExclamation: Exit Sub
    If UBound(strData) <> 42 Then MsgBox "Andere versie Layout retourblok.  Installeer recente versie van GroeneKaart a.u.b.", vbExclamation: Exit Sub
             
    cbSchoon_Click
    For COUNT_TO = 0 To UBound(strTB2)
        Me.mfgLijst.AddItem strTB2(COUNT_TO) & vbTab & strData(COUNT_TO), mfgLijst.Rows - 1
    Next
    
    'Polisnummer
    If Mid(Me.mfgLijst.TextMatrix(2, 0), 1, 7) = "RFF+001" Then
        Me.mebTekstinfo(3).text = Me.mfgLijst.TextMatrix(2, 1)
    End If
    
    'Naam Verzekeringsnemer
    If Mid(Me.mfgLijst.TextMatrix(5, 0), 1, 7) = "PTY+003" Then
        Me.mebTekstinfo(0).text = Me.mfgLijst.TextMatrix(5, 1)
    End If

    'Straat en nr.
    If Mid(Me.mfgLijst.TextMatrix(7, 0), 1, 7) = "ADR+002" Then
        Me.mebTekstinfo(1).text = Me.mfgLijst.TextMatrix(7, 1)
    End If

    'Plaat
    If Mid(Me.mfgLijst.TextMatrix(14, 0), 1, 7) = "RFF+010" Then
        Me.mebTekstinfo(4).text = Me.mfgLijst.TextMatrix(14, 1)
    End If

    'Merk en type
    If Mid(Me.mfgLijst.TextMatrix(10, 0), 1, 7) = "ROD+001" Then
        Me.mebTekstinfo(5).text = Me.mfgLijst.TextMatrix(10, 1)
    End If
    
    Dim Mij As String
    'Verzekeraar
    If Mid(Me.mfgLijst.TextMatrix(3, 0), 1, 7) = "PTY+006" Then
        Mij = Mid(Me.mfgLijst.TextMatrix(3, 1), 2)
    End If
    
    For COUNT_TO = 0 To Me.cmbMaatschappij.ListCount
        If Mid(Me.cmbMaatschappij.List(COUNT_TO), 1, 4) = Mij Then
            Me.cmbMaatschappij.ListIndex = COUNT_TO
            KtrlBox = grafischGKOphalen
            Exit Sub
        End If
    Next
    MsgBox "Afbeelding groene kaart voor verzekeraar " & Mij & " niet beschikbaar.  Eerst volledige versie installeren a.u.b.", vbExclamation
    Exit Sub
    
CancelError:
    Screen.MousePointer = vbNormal

End Sub

Private Sub cbSchoon_Click()

    Dim COUNT_TO As Integer
    
    Me.mfgLijst.Clear
    Me.mfgLijst.Cols = 2
    Me.mfgLijst.Rows = 2
    Me.mfgLijst.ColWidth(0) = 2145
    Me.mfgLijst.ColAlignment(0) = 0
    Me.mfgLijst.ColAlignment(1) = 0
    Me.mfgLijst.ColWidth(1) = 2475
    Me.cmbMaatschappij.ListIndex = 0
    For COUNT_TO = 0 To 5
        Me.mebTekstinfo(COUNT_TO).text = ""
    Next
    Set imgFiguur.Picture = Nothing
    
End Sub

Private Sub cbTogglePrinter_Click()

    Me.cbDocumentPrinterHier.Enabled = Not Me.cbDocumentPrinterHier.Enabled
    
End Sub

Private Sub cbXYwijzigen_Click()

    Dim HetBESTAND As String

    HetBESTAND = LOCATION_COMPANYDATA & "\gkauto.600"
    If Dir(HetBESTAND) = "" Then
        MsgBox HetBESTAND & " is niet aanwezig", vbInformation
    Else
        Shell "notepad.exe " + HetBESTAND, vbNormalFocus
    End If

End Sub

Private Sub cmbMaatschappij_Change()

    cmdWijzigen.Enabled = True

End Sub


Private Sub cmbMaatschappij_Click()
    
    Dim COUNT_TO As Integer

    For COUNT_TO = 0 To mebXYPositie.UBound
        mebXYPositie(COUNT_TO).text = "000/000"
        mebXYPositie(COUNT_TO).Enabled = False
    Next
    cmbFormaatVan.Enabled = False
    cmbFormaatTot.Enabled = False

    cmdWijzigen.Enabled = False
    If cmbMaatschappij.ListIndex = 0 Then
        cmdAfdrukken.Enabled = False
        cmdBewaren.Enabled = False
    Else
        'inladen van gegevens proberen
        'indien mogelijk dan afdruk mogelijk maken
        If LaadMijInfo(cmbMaatschappij.text) Then
            cmdAfdrukken.Enabled = True
            KtrlBox = grafischGKOphalen
        Else
            cmdAfdrukken.Enabled = False
        End If
        cmdWijzigen.Enabled = True
    End If


End Sub


Private Sub cmbMaatschappij_KeyPress(KeyAscii As Integer)

    If Command = "" Then KeyAscii = 0

End Sub


Private Sub cmdAfdrukken_Click()
    
    Dim KtrlBox As Integer
    
    KtrlBox = MsgBox("Breng één exemplaar voorgedrukte groene kaart voor deze verzekeraar in uw printer en druk OK voor vervolg.", vbOKCancel + vbDefaultButton2 + vbExclamation)
    If KtrlBox = vbOK Then
        Screen.MousePointer = vbHourglass
        DrukTekstAf
        Printer.EndDoc
        Printer.ScaleMode = vbTwips
        Screen.MousePointer = vbNormal
    End If
    
End Sub


Private Sub cmdBewaren_Click()

    Me.MousePointer = vbHourglass
    BewaarMijInfo cmbMaatschappij.text
    Me.MousePointer = vbNormal

End Sub


Private Sub cmdHistoriek_Click()
 
    Shell "notepad.exe " + App.path + "\gkauto.log", vbNormalFocus
 
End Sub


Private Sub cmdSluiten_Click()

    Unload Me
    
End Sub


Private Sub cmdUpdate_Click()

    Dim slHier As String
    Dim SQLstring As String
    
    slHier = tbGK(0).text
    If rs.State Then
        rs.Close
    End If
    SQLstring = "SELECT * FROM NL_GroeneKaart WHERE tbSleutel = '" & slHier & "'"
    fnOpenDatabase (SQLstring)
    If rs.RecordCount = 1 Then
        'ok
        For COUNT_TO = 1 To 2
            rs(tbGK(COUNT_TO).Tag).Value = tbGK(COUNT_TO).text
        Next
        rs.Update
     Else
        MsgBox "Gegevens voor '" & slHier & "' niet gevonden", vbInformation
    End If
    rs.Close
   
End Sub

Private Sub cmdWijzigen_Click()

    Dim COUNT_TO As Integer

    If Command <> "/GROENE KAART" Then
        MsgBox "Wijzigen toggle (/GROENE KAART) aanduiden bij opstart programma a.u.b.  Bel 053/21.59.25 of fax 053/78.19.22 voor meer informatie.", vbCritical
        Exit Sub
    End If

    For COUNT_TO = 0 To mebXYPositie.UBound
        mebXYPositie(COUNT_TO).Enabled = True
    Next
    cmbFormaatVan.Enabled = True
    cmbFormaatTot.Enabled = True

End Sub


Private Sub Form_Load()

    Dim flHier As Integer
    Dim Zoekstring As String

    cbDocumentPrinterHier.Clear
    If Printers.Count = 0 Then MsgBox "Installeer eerst een printerdriver a.u.b. (ook al beschikt U nog over geen printer.  Bij blijvende problemen in netwerk gelieve op te starten met het command '/P=FALSE' en uw netwerkbeheerder te informeren)", vbExclamation: End
    For COUNT_TO = 0 To Printers.Count - 1
        Me.cbDocumentPrinterHier.AddItem Printers(COUNT_TO).Port + " " + Printers(COUNT_TO).DeviceName
    Next
    cbDocumentPrinterHier.ListIndex = dokumentPrinterNr

    'On Error GoTo FoutBestand

    Set cnn = New ADODB.Connection
    Set rs = New ADODB.Recordset

    dtpVan.Value = Now
    dtpTot.Value = Now + 60
    cmbSoort.ListIndex = 0
    If Dir(LOCATION_COMPANYDATA & "gkauto.600") = "" Then
    Else
        cmbMaatschappij.Clear
        cmbMaatschappij.AddItem "Eerst maatschapij kiezen a.u.b. !"
        flHier = FreeFile
        Open LOCATION_COMPANYDATA & "\gkauto.600" For Input As flHier
        Do While Not EOF(flHier)
            Line Input #flHier, Zoekstring
            If Left(Zoekstring, 4) = "IDK:" Then
                cmbMaatschappij.AddItem Mid(Zoekstring, 5), cmbMaatschappij.ListCount
            End If
        Loop
        Close flHier
    End If
    cmbMaatschappij.ListIndex = 0
    cbSchoon_Click
    Exit Sub
    
FoutBestand:
    MsgBox LOCATION_COMPANYDATA & "\gkauto.600" + vbCrLf + " is defekt.  Installeer opnieuw of kontakteer R&VSoft !", vbCritical
    cmbMaatschappij.Enabled = False

End Sub



Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    rs.Close
    cnn.Close
    Set rs = Nothing
    Set cnn = Nothing
        
End Sub


Private Sub mebXYPositie_LostFocus(Index As Integer)

    cmdBewaren.Enabled = True

End Sub

Function LaadMijInfo(Sleutel As String) As Integer

    Dim flHier As Integer
    Dim COUNT_TO As Integer
    Dim Zoekstring As String

    LaadMijInfo = False
    If Dir(LOCATION_COMPANYDATA & "\gkauto.600") = "" Then
        MsgBox LOCATION_COMPANYDATA & "\gkauto.600" + vbCrLf + vbCrLf + "is nog niet aanwezig."
        Exit Function
    Else
        flHier = FreeFile
        Open LOCATION_COMPANYDATA & "\gkauto.600" For Input As flHier
        Do While Not EOF(flHier)
            Line Input #flHier, Zoekstring
            If Zoekstring = "IDK:" + Sleutel Then
                For COUNT_TO = 0 To mebXYPositie.UBound
                    Line Input #flHier, Zoekstring
                    mebXYPositie(COUNT_TO).text = Zoekstring
                Next
                LaadMijInfo = True
                Exit Do
            End If
        Loop
        Line Input #flHier, Zoekstring
            cmbFormaatVan.ListIndex = Val(Zoekstring)
        Line Input #flHier, Zoekstring
            cmbFormaatTot.ListIndex = Val(Zoekstring)
        Close flHier
    End If

End Function


Sub BewaarMijInfo(Sleutel As String)

    Dim flHier As Integer
    Dim COUNT_TO As Integer
    Dim FlBackup As Integer
    Dim Zoekstring As String

    flHier = FreeFile
    If Dir(App.path + "\gkauto.600") = "" Then
    Else
        'Eerst vorige definitie zoeken en verwijderen
        Open App.path + "\gkauto.600" For Input As flHier
        FlBackup = FreeFile
        Open App.path + "\gkauto.$$$" For Output As FlBackup
        Do While Not EOF(flHier)
            Line Input #flHier, Zoekstring
            If Zoekstring = "IDK:" + Sleutel Then
                For COUNT_TO = 0 To mebXYPositie.UBound
                    Line Input #flHier, Zoekstring
                Next
                Line Input #flHier, Zoekstring 'formaatvan
                Line Input #flHier, Zoekstring 'formaattot
            Else
                Print #FlBackup, Zoekstring + vbCrLf;
                For COUNT_TO = 0 To mebXYPositie.UBound
                    Line Input #flHier, Zoekstring
                    Print #FlBackup, Zoekstring + vbCrLf;
                Next
                Line Input #flHier, Zoekstring  'formaatvan
                Print #FlBackup, Zoekstring + vbCrLf;
                Line Input #flHier, Zoekstring  'formaattot
                Print #FlBackup, Zoekstring + vbCrLf;
            End If
        Loop
        Close flHier
        Close FlBackup
        Kill App.path + "\gkauto.600"
        If Dir(App.path + "\gkauto.$$$") = "" Then
        Else
            Name App.path + "\gkauto.$$$" As App.path + "\gkauto.600"
        End If
    End If

    Open App.path + "\gkauto.600" For Append As flHier
    Print #flHier, "IDK:" + Sleutel + vbCrLf;
    For COUNT_TO = 0 To mebXYPositie.UBound
        Print #flHier, mebXYPositie(COUNT_TO).text + vbCrLf;
    Next
    Print #flHier, Format(cmbFormaatVan.ListIndex) + vbCrLf;
    Print #flHier, Format(cmbFormaatTot.ListIndex) + vbCrLf;
    Close flHier

End Sub


Private Sub cmdAfdruk_Click()
    
    Dim KtrlBox As Integer
    Dim COUNT_TO As Integer
        
    If Me.cmbMaatschappij.ListIndex = 0 Then MsgBox "Verzekeraar aanduiden a.u.b.", vbExclamation: Exit Sub
    For COUNT_TO = 0 To 5
        If COUNT_TO = 2 Then
        Else
            If Me.mebTekstinfo(COUNT_TO).text = "" Then MsgBox "Eerst vervolledigen a.u.b.", vbInformation: Exit For
        End If
    Next
    KtrlBox = MsgBox("Sommige verzekeraars gaan mogelijk niet akkoord met deze werkwijze.  U dient dit zelf onder controle te houden onder eigen verantwoordelijkheid." & vbCrLf & vbCrLf & "Vervolg met afdrukken ?", vbDefaultButton2 + vbYesNo + vbQuestion)
    If KtrlBox = vbYes Then
        PrintAnywhere imgFiguur, Printer
    End If
  
End Sub


Private Sub cmdBestand_Click()
   
    On Error GoTo CancelError
    Mim.Teken.fileName = ""
    Mim.Teken.CancelError = True
    Mim.Teken.Filter = _
        "Alle Figuurbestanden|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur;*.tif" & _
        "|JPEG bestanden (*.jpg)|*.jpg" & _
        "|GIF bestanden (*.gif)|*.gif" & _
        "|BITMAP bestanden (*.bmp;*.dib)|*.bmp;*.dib" & _
        "|META bestanden (*.wmf, *.emf)|*.wmf;*.emf" & _
        "|CURSOR bestanden (*.ico, *.cur)|*.ico;*.cur"
                
    Mim.Teken.ShowOpen
    imgFiguur.Picture = LoadPicture(Mim.Teken.fileName)
    cmdBewaar.Enabled = True
    cmdBewaar.SetFocus
    BestandsNaamFiguur = Mim.Teken.fileName
    Exit Sub
    
CancelError:
    Screen.MousePointer = vbNormal

End Sub


Private Sub CmdBewaar_Click()

    If rs.State Then
    Else
        fnOpenDatabase ("NL_GroeneKaart")
    End If
    
BewaarAfbeelding:
    KtrlBox = MsgBox("Bewaren als nieuw", vbQuestion + vbYesNoCancel + vbDefaultButton3)
    If KtrlBox = vbCancel Then
        GoTo BewaarAfbeelding
    ElseIf KtrlBox = vbNo Then
        'Eerst zoeken...
        Stop
    Else
        rs.AddNew
    End If
    For COUNT_TO = 0 To 2
        rs(tbGK(COUNT_TO).Tag).Value = tbGK(COUNT_TO).text
    Next
    DoEvents
    FileToBlob rs("tbobject"), BestandsNaamFiguur
    rs.Update
    cmdBestand.SetFocus
    cmdBewaar.Enabled = False

End Sub


Private Sub cmdBodem_Click()

    cmdBewaar.Enabled = False
    Schoon
    
    On Error GoTo ErrorMOVELAST
    If rs.State Then
    Else
        fnOpenDatabase ("NL_GroeneKaart")
    End If
    rs.MoveLast
    GKOphalen
    
ErrorMOVELAST:
        
End Sub


Private Sub cmdScan_Click()
   
    cmdBewaar.Enabled = False

    If rs.State Then
    Else
        fnOpenDatabase ("NL_GroeneKaart")
    End If
    
    Screen.MousePointer = vbHourglass
    'ImgScan1.ShowSetupBeforeScan = True
    'ImgScan1.ShowSelectScanner
        
    'If ImgScan1.StatusCode = 0 Then
        'scanner werd gekozen
    '    ImgScan1.OpenScanner
    '    If ImgScan1.StatusCode = 0 Then
    '        'scanner kan geopend worden
    '        'TIFF, JPG_File, BMP_Bitmap, AWD_MicrosoftFax
    '        '  1 ,       6 ,         3 ,                2
    '
    '        ImgScan1.FileType = JPG_File
    '        ImgScan1.Image = App.path & "\marscan.jpg"
    '        ImgScan1.PageOption = PromptToCreateNewFile
    '
    '        ImgScan1.StartScan
    '        If ImgScan1.StatusCode = 0 Then
    '            'scan met succes
    '            'opslaan en afbeelden
    '            imgFiguur.Picture = LoadPicture(ImgScan1.Image)
    '            'BewaarAfbeelding
    '            rs.AddNew
    '            rs("tbDatum") = Now
    '            rs("tbOmschrijving") = "Een testafbeelding " & Now
    '            DoEvents
    '            FileToBlob rs("tbobject"), ImgScan1.Image
    '            rs.Update
    '        End If
    '    End If
    '    ImgScan1.CloseScanner
    'End If
    Screen.MousePointer = vbNormal
    
End Sub


Private Sub cmdScanNieuweAfbeelding_Click()

    cmdBewaar.Enabled = False
    
    'ImgScan1.ShowScanNew
    
End Sub


Private Sub cmdScanOpties_Click()

    cmdBewaar.Enabled = False

    'ImgScan1.ShowScanPreferences
    
End Sub


Private Sub cmdTop_Click()

    Schoon
    cmdBewaar.Enabled = False
    
    On Local Error GoTo ErrorMOVEFIRST
    If rs.State Then
    Else
        fnOpenDatabase ("NL_GroeneKaart")
    End If
    rs.MoveFirst
    BlobToFile rs("tbobject"), App.path & "\marscan.jpg"
    For COUNT_TO = 0 To 2
        tbGK(COUNT_TO).text = rs(tbGK(COUNT_TO).Tag).Value
    Next
    On Error Resume Next
    imgFiguur.Picture = LoadPicture(App.path & "\marscan.jpg")
    If Err Then MsgBox Error
    
ErrorMOVEFIRST:

End Sub


Private Sub cmdVolgende_Click()

    cmdBewaar.Enabled = False
    Schoon
    
    On Error GoTo ErrorMOVENEXT
    If rs.State Then
    Else
        cmdBodem_Click
        Exit Sub
    End If
    rs.MoveNext
    If rs.EOF Then
        cmdBodem_Click
        Exit Sub
    Else
        GKOphalen
    End If
    
ErrorMOVENEXT:

End Sub


Private Sub cmdVorige_Click()

    Schoon
    cmdBewaar.Enabled = False
        
    On Local Error GoTo ErrorBOF
    If rs.State Then
    Else
        cmdTop_Click
        Exit Sub
    End If
    rs.MovePrevious
    If rs.BOF Then
        cmdTop_Click
        Exit Sub
    Else
        GKOphalen
    End If
    
ErrorBOF:
    
End Sub


Function fnOpenDatabase(SQLBevel As String)

    If Dir(LOCATION_COMPANYDATA & "ASScan.mdb") = "" Then
        If Not CopyFile(App.path, LOCATION_COMPANYDATA, "ASScan.mdb") Then
            MsgBox "ASScan.mdb is noodzakelijk.", vbExclamation
            Unload Me
            Exit Function
        End If
    End If
    
    On Error Resume Next
    cnn.Close

    cnn.Open ADOJET_PROVIDER & _
        "Data Source=" & LOCATION_COMPANYDATA & _
        "ASScan.mdb;" & _
        "Persist Security Info=False"
    rs.CursorLocation = adUseClient
    rs.Open SQLBevel, cnn, adOpenDynamic, adLockOptimistic

End Function


' Copy a BLOB field's contents to a binary file.
Function BlobToFile(fld As ADODB.Field, fileName As String, _
    Optional ChunkSize As Long = 8192)
    Dim fnum As Integer, bytesLeft As Long, bytes As Long
    Dim tmp() As Byte
    
    ' Raise an error if the field doesn't support GetChunk.
    If (fld.Attributes And adFldLong) = 0 Then
        Err.Raise 1001, , "Field doesn't support the GetChunk method."
    End If
    ' Open the file;, delete it firstoverwrite it if necessary.' Delete the
    ' file if it exists already, then create a new one.
    If Dir$(fileName) <> "" Then Kill fileName
    
    fnum = FreeFile
    Open fileName For Binary As fnum
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
Function FileToBlob(fld As ADODB.Field, fileName As String, _
    Optional ChunkSize As Long = 8192)
    
    Dim fnum As Integer, bytesLeft As Long, bytes As Long
    Dim tmp() As Byte
    
    ' Raise an error if the field doesn't support GetChunk.
    If (fld.Attributes And adFldLong) = 0 Then
        Err.Raise 1001, , "Field doesn't support the GetChunk method."
    End If
    ' Open the file; raise an error if the file doesn't exist.
    If Dir$(fileName) = "" Then Err.Raise 53, , "File not found"
    
    fnum = FreeFile
    Open fileName For Binary As fnum
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


Sub PrintAnywhere(Src As Object, Dest As Object)
    
    Err = 0
    On Error Resume Next
    If Dest Is Printer Then
        Set Printer = Printers(dokumentPrinterNr)
        On Error Resume Next
        'Printer.PaperBin = LaadTekst(App.Title, "dokumentPRINTER")
        'If Printer.Orientation = vbPRORLandscape Then
        '    Printer.Orientation = vbPRORPortrait
        '    DoEvents
        'End If
        Select Case Me.cbKwaliteit.ListIndex
            Case 0
                Printer.PrintQuality = vbPRPQDraft
                Refresh
            Case 1
                Printer.PrintQuality = vbPRPQLow
                Refresh
            Case 2
                Printer.PrintQuality = vbPRPQMedium
                Refresh
            Case 3
                Printer.PrintQuality = vbPRPQHigh
                Refresh
        End Select
        If Me.cbKleur.Value = vbChecked Then
            Printer.ColorMode = vbPRCMColor
        Else
            Printer.ColorMode = vbPRCMMonochrome
        End If
        DoEvents
        Printer.PaintPicture Src.Picture, 0, 0, Printer.Width - 500, Printer.Height - 500
        DoEvents
        If Left(cmbMaatschappij.text, 4) = tbGK(0).text Then
            DrukTekstAf
        End If
        Printer.EndDoc
        Printer.ScaleMode = vbTwips
    Else
        Dest.PaintPicture Src.Picture, Dest.Width / 2, Dest.Height / 2
    End If
    If Err Then MsgBox Error, vbInformation
   
End Sub


Sub ResizePictureBoxToImage(pic As PictureBox, twipWd _
  As Integer, twipHt As Integer)
 
    ' This code assumes that all units are in twips.  If
    ' not, you must convert it to twips before calling
    ' this routine.  This also assumes that the image
    ' was blt'ed to 0,0.
    Dim BorderHt As Integer, BorderWd As Integer
    BorderWd = pic.Width - pic.ScaleWidth
    BorderHt = pic.Height - pic.ScaleHeight
    pic.Move pic.Left, pic.Top, twipWd + BorderWd, _
    twipHt + BorderHt

End Sub


Sub Schoon()

    For COUNT_TO = 0 To 2
        tbGK(COUNT_TO).text = ""
    Next
    Me.Refresh
    
End Sub
Function grafischGKOphalen() As Boolean

    Dim slHier As String
    Dim SQLstring As String

        cbKwaliteit.ListIndex = 2
        slHier = Left(cmbMaatschappij.text, 4)
        If UCase(slHier) = "EERS" Then Exit Function
        If rs.State Then
            rs.Close
        End If
        SQLstring = "SELECT * FROM NL_GroeneKaart WHERE tbSleutel = '" & slHier & "'"
        fnOpenDatabase (SQLstring)
        If rs.RecordCount = 1 Then
            'ok
            GKOphalen
            GKIsKLAAR = True
        Else
            MsgBox "Gegevens voor '" & slHier & "' niet gevonden", vbInformation
        End If
        rs.Close
        
End Function


Function GKOphalen()
    
    Screen.MousePointer = vbHourglass
    BlobToFile rs("tbobject"), App.path & "\marscan.jpg"
    For COUNT_TO = 0 To 2
        tbGK(COUNT_TO).text = rs(tbGK(COUNT_TO).Tag).Value
    Next
    On Error Resume Next
    imgFiguur.Picture = LoadPicture(App.path & "\marscan.jpg")
      
    If Err Then MsgBox Error
    Screen.MousePointer = vbNormal
    GKIsKLAAR = False

End Function


Function DrukTekstAf()

    Dim xVerlies As Single
    Dim yVerlies As Single
    Dim COUNT_TO As Integer
    Dim flHier As Integer

    On Local Error GoTo TEKSTERROR
    
    Me.MousePointer = vbHourglass
    flHier = FreeFile
    Open LOCATION_COMPANYDATA & "gkauto.log" For Append As flHier

    xVerlies = Left(mebXYPositie(9), 3)   'Horizontaal
    yVerlies = Right(mebXYPositie(9), 3)  'Vertikaal
    Printer.ScaleMode = vbMillimeters
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    
    Print #flHier, vbCrLf; "Voorlopige groene kaart, " + cmbMaatschappij.text + " op " + Format(Now) + vbCrLf;
    For COUNT_TO = 0 To 5
        Printer.CurrentX = Left(mebXYPositie(COUNT_TO), 3) - xVerlies 'Horizontaal
        Printer.CurrentY = Right(mebXYPositie(COUNT_TO), 3) - yVerlies 'Vertikaal
        Printer.Print mebTekstinfo(COUNT_TO);
        Print #flHier, mebTekstinfo(COUNT_TO); vbCrLf;
    Next
    
    Printer.CurrentX = Left(mebXYPositie(6).text, 3) - xVerlies 'Horizontaal
    Printer.CurrentY = Right(mebXYPositie(6).text, 3) - yVerlies 'Vertikaal
    Printer.Print Left(cmbSoort.text, 1);
    Print #flHier, cmbSoort; vbCrLf;
    
    Printer.CurrentX = Left(mebXYPositie(7).text, 3) - xVerlies 'Horizontaal
    Printer.CurrentY = Right(mebXYPositie(7).text, 3) - yVerlies 'Vertikaal
    Printer.Print Format(dtpVan, cmbFormaatVan);
    Print #flHier, Format(dtpVan, cmbFormaatVan); vbCrLf;
    
    Printer.CurrentX = Left(mebXYPositie(8).text, 3) - xVerlies 'Horizontaal
    Printer.CurrentY = Right(mebXYPositie(8).text, 3) - yVerlies 'Vertikaal
    Printer.Print Format(dtpTot, cmbFormaatTot);
    Print #flHier, Format(dtpTot, cmbFormaatTot); vbCrLf;
    Close flHier
    
TEKSTERROR:
If Err Then
    MsgBox Err.Description
    Close flHier
Else
    Close flHier
    
End If
Me.MousePointer = vbNormal

End Function
