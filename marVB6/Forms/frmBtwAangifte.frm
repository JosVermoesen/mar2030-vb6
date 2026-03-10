VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form frmBTWAangifte 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BTW Aangifte België (Model EDIFACT X400)"
   ClientHeight    =   7935
   ClientLeft      =   990
   ClientTop       =   1080
   ClientWidth     =   10530
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7935
   ScaleWidth      =   10530
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cbIntervat2025 
      Caption         =   "XML 04/2025"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9240
      Picture         =   "frmBtwAangifte.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   205
      Top             =   0
      Width           =   1155
   End
   Begin VB.TextBox tbmailbtw 
      Height          =   285
      Left            =   7560
      TabIndex        =   204
      Text            =   "info@rv.be"
      ToolTipText     =   "doormailen ofwel naar vsoft ofwel naar uw raadgever voor verder gevolg"
      Top             =   720
      Width           =   2895
   End
   Begin VB.CommandButton cbIntervat2008 
      Caption         =   "XML 04/2009"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   7560
      Picture         =   "frmBtwAangifte.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   201
      Top             =   0
      Visible         =   0   'False
      Width           =   1155
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6675
      Left            =   2160
      TabIndex        =   14
      Top             =   1200
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   11774
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Model 2003"
      TabPicture(0)   =   "frmBtwAangifte.frx":0884
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "lblInfoTekst"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Model 1999"
      TabPicture(1)   =   "frmBtwAangifte.frx":08A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblEURVak(62)"
      Tab(1).Control(1)=   "Label1(23)"
      Tab(1).Control(2)=   "lblVakCaption(87)"
      Tab(1).Control(3)=   "lblVakCaption(86)"
      Tab(1).Control(4)=   "lblVakCaption(85)"
      Tab(1).Control(5)=   "lblVakCaption(84)"
      Tab(1).Control(6)=   "lblVakCaption(83)"
      Tab(1).Control(7)=   "lblVakCaption(82)"
      Tab(1).Control(8)=   "lblVakCaption(81)"
      Tab(1).Control(9)=   "lblVakCaption(64)"
      Tab(1).Control(10)=   "lblVakCaption(63)"
      Tab(1).Control(11)=   "lblVakCaption(59)"
      Tab(1).Control(12)=   "lblVakCaption(57)"
      Tab(1).Control(13)=   "lblVakCaption(56)"
      Tab(1).Control(14)=   "lblVakCaption(55)"
      Tab(1).Control(15)=   "lblVakCaption(54)"
      Tab(1).Control(16)=   "lblVakCaption(49)"
      Tab(1).Control(17)=   "lblVakCaption(48)"
      Tab(1).Control(18)=   "lblVakCaption(47)"
      Tab(1).Control(19)=   "lblVakCaption(46)"
      Tab(1).Control(20)=   "lblVakCaption(45)"
      Tab(1).Control(21)=   "lblVakCaption(3)"
      Tab(1).Control(22)=   "lblVakCaption(2)"
      Tab(1).Control(23)=   "lblVakCaption(1)"
      Tab(1).Control(24)=   "lblVakCaption(0)"
      Tab(1).Control(25)=   "lblBEFVak(0)"
      Tab(1).Control(26)=   "lblBEFVak(1)"
      Tab(1).Control(27)=   "lblBEFVak(2)"
      Tab(1).Control(28)=   "lblBEFVak(3)"
      Tab(1).Control(29)=   "lblBEFVak(45)"
      Tab(1).Control(30)=   "lblBEFVak(46)"
      Tab(1).Control(31)=   "lblBEFVak(47)"
      Tab(1).Control(32)=   "lblBEFVak(48)"
      Tab(1).Control(33)=   "lblBEFVak(49)"
      Tab(1).Control(34)=   "lblBEFVak(54)"
      Tab(1).Control(35)=   "lblBEFVak(55)"
      Tab(1).Control(36)=   "lblBEFVak(56)"
      Tab(1).Control(37)=   "lblBEFVak(57)"
      Tab(1).Control(38)=   "lblBEFVak(59)"
      Tab(1).Control(39)=   "lblBEFVak(63)"
      Tab(1).Control(40)=   "lblBEFVak(64)"
      Tab(1).Control(41)=   "lblBEFVak(81)"
      Tab(1).Control(42)=   "lblBEFVak(82)"
      Tab(1).Control(43)=   "lblBEFVak(83)"
      Tab(1).Control(44)=   "lblBEFVak(84)"
      Tab(1).Control(45)=   "lblBEFVak(85)"
      Tab(1).Control(46)=   "lblBEFVak(86)"
      Tab(1).Control(47)=   "lblBEFVak(87)"
      Tab(1).Control(48)=   "lblBEFVak(88)"
      Tab(1).Control(49)=   "lblVakCaption(61)"
      Tab(1).Control(50)=   "lblBEFVak(61)"
      Tab(1).Control(51)=   "lblVakCaption(65)"
      Tab(1).Control(52)=   "lblBEFVak(65)"
      Tab(1).Control(53)=   "lblBEFVakXX(0)"
      Tab(1).Control(54)=   "lblVakXXCaption(0)"
      Tab(1).Control(55)=   "lblBEFVak(62)"
      Tab(1).Control(56)=   "lblVakCaption(62)"
      Tab(1).Control(57)=   "lblBEFVak(66)"
      Tab(1).Control(58)=   "lblVakCaption(66)"
      Tab(1).Control(59)=   "lblBEFVakYY(0)"
      Tab(1).Control(60)=   "lblVakYYCaption(0)"
      Tab(1).Control(61)=   "lblBEF(0)"
      Tab(1).Control(62)=   "lblBEF(1)"
      Tab(1).Control(63)=   "lblBEFVak(71)"
      Tab(1).Control(64)=   "lblBEFVak(72)"
      Tab(1).Control(65)=   "lblBEFVak(91)"
      Tab(1).Control(66)=   "lblVakCaption(71)"
      Tab(1).Control(67)=   "lblVakCaption(72)"
      Tab(1).Control(68)=   "lblVakCaption(91)"
      Tab(1).Control(69)=   "lblEURVak(0)"
      Tab(1).Control(70)=   "lblBEF(2)"
      Tab(1).Control(71)=   "lblEURVak(1)"
      Tab(1).Control(72)=   "lblEURVak(2)"
      Tab(1).Control(73)=   "lblEURVak(3)"
      Tab(1).Control(74)=   "lblEURVak(45)"
      Tab(1).Control(75)=   "lblEURVak(46)"
      Tab(1).Control(76)=   "lblEURVak(47)"
      Tab(1).Control(77)=   "lblEURVak(48)"
      Tab(1).Control(78)=   "lblEURVak(49)"
      Tab(1).Control(79)=   "lblEURVak(81)"
      Tab(1).Control(80)=   "lblEURVak(82)"
      Tab(1).Control(81)=   "lblEURVak(83)"
      Tab(1).Control(82)=   "lblEURVak(84)"
      Tab(1).Control(83)=   "lblEURVak(85)"
      Tab(1).Control(84)=   "lblEURVak(86)"
      Tab(1).Control(85)=   "lblEURVak(87)"
      Tab(1).Control(86)=   "Line1"
      Tab(1).Control(87)=   "lblEURVak(54)"
      Tab(1).Control(88)=   "lblEURVak(55)"
      Tab(1).Control(89)=   "lblEURVak(56)"
      Tab(1).Control(90)=   "lblEURVak(57)"
      Tab(1).Control(91)=   "lblEURVak(61)"
      Tab(1).Control(92)=   "lblEURVak(63)"
      Tab(1).Control(93)=   "lblEURVak(65)"
      Tab(1).Control(94)=   "lblEURVakXX(0)"
      Tab(1).Control(95)=   "lblEURVak(59)"
      Tab(1).Control(96)=   "lblEURVak(64)"
      Tab(1).Control(97)=   "lblEURVak(66)"
      Tab(1).Control(98)=   "lblEURVakYY(0)"
      Tab(1).Control(99)=   "lblEURVak(71)"
      Tab(1).Control(100)=   "lblEURVak(72)"
      Tab(1).Control(101)=   "lblEURVak(91)"
      Tab(1).Control(102)=   "lblBEF(3)"
      Tab(1).ControlCount=   103
      TabCaption(2)   =   "Intervat 2008"
      TabPicture(2)   =   "frmBtwAangifte.frx":08BC
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Command1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cbVergrendel"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "RichTextBox1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cbAanvraagTerugbetaling"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cbAanvraagBetaalformulieren"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "Intervat 2025"
      TabPicture(3)   =   "frmBtwAangifte.frx":08D8
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cbXml2025"
      Tab(3).Control(1)=   "cbPayment2025"
      Tab(3).Control(2)=   "cbRestitution2025"
      Tab(3).Control(3)=   "RichTextBox2"
      Tab(3).ControlCount=   4
      Begin VB.CommandButton cbXml2025 
         Caption         =   "XML Bestand"
         Height          =   495
         Left            =   -68400
         TabIndex        =   209
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CheckBox cbPayment2025 
         Alignment       =   1  'Right Justify
         Caption         =   "Aanvraag van Betaalformulieren"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -69480
         TabIndex        =   208
         Top             =   480
         Width           =   2655
      End
      Begin VB.CheckBox cbRestitution2025 
         Caption         =   "Aanvraag om Terugbetaling"
         Height          =   195
         Left            =   -74880
         TabIndex        =   206
         Top             =   480
         Width           =   2895
      End
      Begin VB.CheckBox cbAanvraagBetaalformulieren 
         Alignment       =   1  'Right Justify
         Caption         =   "Aanvraag van Betaalformulieren"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5400
         TabIndex        =   203
         Top             =   480
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CheckBox cbAanvraagTerugbetaling 
         Caption         =   "Aanvraag om Terugbetaling"
         Height          =   195
         Left            =   120
         TabIndex        =   202
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   5775
         Left            =   120
         TabIndex        =   197
         Top             =   780
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   10186
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmBtwAangifte.frx":08F4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Height          =   4935
         Left            =   -74820
         TabIndex        =   116
         ToolTipText     =   "In ontwikkeling per 18/11/2003"
         Top             =   780
         Width           =   7335
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   62
            Left            =   3900
            TabIndex        =   196
            Tag             =   "v045"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   2520
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   17
            Left            =   5160
            TabIndex        =   195
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   4230
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   72
            Left            =   5160
            TabIndex        =   194
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   3900
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   71
            Left            =   5160
            TabIndex        =   193
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   3630
            Width           =   1200
         End
         Begin VB.Label lblBEFVakYY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   5160
            TabIndex        =   192
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   3300
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   66
            Left            =   5160
            TabIndex        =   191
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   3030
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   62
            Left            =   5160
            TabIndex        =   190
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   2490
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   5160
            TabIndex        =   189
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   1890
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   65
            Left            =   5160
            TabIndex        =   188
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   1620
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   61
            Left            =   5160
            TabIndex        =   187
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   1080
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   64
            Left            =   5160
            TabIndex        =   186
            Tag             =   "v101"
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   2760
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   63
            Left            =   5160
            TabIndex        =   185
            Tag             =   "v100"
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   1350
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   59
            Left            =   5160
            TabIndex        =   184
            Tag             =   "v045"
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   2220
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   57
            Left            =   5160
            TabIndex        =   183
            Tag             =   "v044"
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   810
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   56
            Left            =   5160
            TabIndex        =   182
            Tag             =   "v043"
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   540
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   55
            Left            =   5160
            TabIndex        =   181
            Tag             =   "v042"
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   300
            Width           =   1200
         End
         Begin VB.Label lblARvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   54
            Left            =   5160
            TabIndex        =   180
            Tag             =   "v064"
            ToolTipText     =   "Volgens stand algemene rekening"
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   72
            Left            =   3900
            TabIndex        =   179
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   3900
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   71
            Left            =   3900
            TabIndex        =   178
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   3630
            Width           =   1200
         End
         Begin VB.Label lblEURVakYY 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3900
            TabIndex        =   177
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   3300
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   66
            Left            =   3900
            TabIndex        =   176
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   3030
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   64
            Left            =   3900
            TabIndex        =   175
            Tag             =   "v101"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   2730
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   59
            Left            =   3900
            TabIndex        =   174
            Tag             =   "v045"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   2220
            Width           =   1200
         End
         Begin VB.Label lblEURVakXX 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3900
            TabIndex        =   173
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   1890
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   65
            Left            =   3900
            TabIndex        =   172
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   1620
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   63
            Left            =   3900
            TabIndex        =   171
            Tag             =   "v100"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   1350
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   61
            Left            =   3900
            TabIndex        =   170
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   1080
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   57
            Left            =   3900
            TabIndex        =   169
            Tag             =   "v044"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   810
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   56
            Left            =   3900
            TabIndex        =   168
            Tag             =   "v043"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   540
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   55
            Left            =   3900
            TabIndex        =   167
            Tag             =   "v042"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   270
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   54
            Left            =   3900
            TabIndex        =   166
            Tag             =   "v064"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 91"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   33
            Left            =   3240
            TabIndex        =   165
            Top             =   4230
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 72"
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
            Height          =   255
            Index           =   32
            Left            =   3240
            TabIndex        =   164
            Top             =   3900
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 71"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   31
            Left            =   3240
            TabIndex        =   163
            Top             =   3630
            Width           =   645
         End
         Begin VB.Label lblVakYYCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak YY"
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
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   162
            Top             =   3300
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 66"
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
            Height          =   255
            Index           =   30
            Left            =   3240
            TabIndex        =   161
            Top             =   3030
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 62"
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
            Height          =   255
            Index           =   29
            Left            =   3240
            TabIndex        =   160
            Top             =   2490
            Width           =   645
         End
         Begin VB.Label lblVakXXCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak XX"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   159
            Top             =   1890
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 65"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   28
            Left            =   3240
            TabIndex        =   158
            Top             =   1620
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 61"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   27
            Left            =   3240
            TabIndex        =   157
            Top             =   1080
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 54"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   26
            Left            =   3240
            TabIndex        =   156
            Top             =   0
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 55"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   25
            Left            =   3240
            TabIndex        =   155
            Top             =   270
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 56"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   24
            Left            =   3240
            TabIndex        =   154
            Top             =   540
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 57"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   23
            Left            =   3240
            TabIndex        =   153
            Top             =   810
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 59"
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
            Height          =   255
            Index           =   22
            Left            =   3240
            TabIndex        =   152
            Top             =   2220
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 63"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   21
            Left            =   3240
            TabIndex        =   151
            Top             =   1350
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0000FF00&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 64"
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
            Height          =   255
            Index           =   20
            Left            =   3240
            TabIndex        =   150
            Top             =   2760
            Width           =   645
         End
         Begin VB.Line Line2 
            X1              =   3180
            X2              =   3180
            Y1              =   0
            Y2              =   4710
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   87
            Left            =   660
            TabIndex        =   149
            Tag             =   "v053"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   4470
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   86
            Left            =   660
            TabIndex        =   148
            Tag             =   "v052"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   4200
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   85
            Left            =   660
            TabIndex        =   147
            Tag             =   "v051"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   3870
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   84
            Left            =   660
            TabIndex        =   146
            Tag             =   "v050"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   3600
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   83
            Left            =   660
            TabIndex        =   145
            Tag             =   "v048"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   3270
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   82
            Left            =   660
            TabIndex        =   144
            Tag             =   "v047"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   3000
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   81
            Left            =   660
            TabIndex        =   143
            Tag             =   "v046"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   2730
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   49
            Left            =   660
            TabIndex        =   142
            Tag             =   "v063"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   2400
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   48
            Left            =   660
            TabIndex        =   141
            Tag             =   "v062"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   2100
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   47
            Left            =   660
            TabIndex        =   140
            Tag             =   "v061"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   1800
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   46
            Left            =   660
            TabIndex        =   139
            Tag             =   "v060"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   1500
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   45
            Left            =   660
            TabIndex        =   138
            Tag             =   "v059"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   1200
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   660
            TabIndex        =   137
            Tag             =   "v058"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   870
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   660
            TabIndex        =   136
            Tag             =   "v057"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   660
            TabIndex        =   135
            Tag             =   "v056"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   330
            Width           =   1200
         End
         Begin VB.Label lblEvak 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   660
            TabIndex        =   134
            Tag             =   "v055"
            ToolTipText     =   "Volgens aankoop- en verkoopboeken"
            Top             =   0
            Width           =   1200
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 00"
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
            Height          =   255
            Index           =   19
            Left            =   0
            TabIndex        =   133
            Top             =   0
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 01"
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
            Height          =   255
            Index           =   18
            Left            =   0
            TabIndex        =   132
            Top             =   330
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 02"
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
            Height          =   255
            Index           =   17
            Left            =   0
            TabIndex        =   131
            Top             =   600
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 03"
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
            Height          =   255
            Index           =   16
            Left            =   0
            TabIndex        =   130
            Top             =   870
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 45"
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
            Height          =   255
            Index           =   15
            Left            =   0
            TabIndex        =   129
            Top             =   1200
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 46"
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
            Height          =   255
            Index           =   14
            Left            =   0
            TabIndex        =   128
            Top             =   1500
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 47"
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
            Height          =   255
            Index           =   13
            Left            =   0
            TabIndex        =   127
            Top             =   1800
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 48"
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
            Height          =   255
            Index           =   12
            Left            =   0
            TabIndex        =   126
            Top             =   2100
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 49"
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
            Height          =   255
            Index           =   11
            Left            =   0
            TabIndex        =   125
            Top             =   2400
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 81"
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
            Height          =   255
            Index           =   10
            Left            =   0
            TabIndex        =   124
            Top             =   2730
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 82"
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
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   123
            Top             =   3000
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 83"
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
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   122
            Top             =   3270
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 84"
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
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   121
            Top             =   3600
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 85"
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
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   120
            Top             =   3870
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 86"
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
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   119
            Top             =   4200
            Width           =   645
         End
         Begin VB.Label lblVakCaption 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Vak 87"
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
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   118
            Top             =   4470
            Width           =   645
         End
      End
      Begin RichTextLib.RichTextBox RichTextBox2 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   207
         Top             =   720
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   8916
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmBtwAangifte.frx":0981
      End
      Begin VB.CheckBox cbVergrendel 
         Caption         =   "Vergrendeld"
         Height          =   195
         Left            =   120
         TabIndex        =   198
         Top             =   5940
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "XML Bestand"
         Height          =   495
         Left            =   6480
         TabIndex        =   199
         Top             =   5940
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   62
         Left            =   -69720
         TabIndex        =   200
         Tag             =   "v045"
         Top             =   3300
         Width           =   1200
      End
      Begin VB.Label lblInfoTekst 
         Caption         =   $"frmBtwAangifte.frx":0A0E
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   -74940
         TabIndex        =   117
         Top             =   5880
         Width           =   7275
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 88 !"
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
         Height          =   255
         Index           =   23
         Left            =   -70380
         TabIndex        =   115
         Top             =   5430
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 87"
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
         Height          =   255
         Index           =   87
         Left            =   -73650
         TabIndex        =   114
         Top             =   5340
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 86"
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
         Height          =   255
         Index           =   86
         Left            =   -73650
         TabIndex        =   113
         Top             =   5070
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 85"
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
         Height          =   255
         Index           =   85
         Left            =   -73650
         TabIndex        =   112
         Top             =   4740
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 84"
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
         Height          =   255
         Index           =   84
         Left            =   -73650
         TabIndex        =   111
         Top             =   4470
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 83"
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
         Height          =   255
         Index           =   83
         Left            =   -73650
         TabIndex        =   110
         Top             =   4140
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 82"
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
         Height          =   255
         Index           =   82
         Left            =   -73650
         TabIndex        =   109
         Top             =   3870
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 81"
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
         Height          =   255
         Index           =   81
         Left            =   -73650
         TabIndex        =   108
         Top             =   3600
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 64"
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
         Height          =   255
         Index           =   64
         Left            =   -70380
         TabIndex        =   107
         Top             =   3630
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 63"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   63
         Left            =   -70380
         TabIndex        =   106
         Top             =   2220
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 59"
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
         Height          =   255
         Index           =   59
         Left            =   -70380
         TabIndex        =   105
         Top             =   3090
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 57"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   57
         Left            =   -70380
         TabIndex        =   104
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 56"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   56
         Left            =   -70380
         TabIndex        =   103
         Top             =   1410
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 55"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   55
         Left            =   -70380
         TabIndex        =   102
         Top             =   1140
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 54"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   54
         Left            =   -70380
         TabIndex        =   101
         Top             =   870
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 49"
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
         Height          =   255
         Index           =   49
         Left            =   -73650
         TabIndex        =   100
         Top             =   3270
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 48"
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
         Height          =   255
         Index           =   48
         Left            =   -73650
         TabIndex        =   99
         Top             =   3000
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 47"
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
         Height          =   255
         Index           =   47
         Left            =   -73650
         TabIndex        =   98
         Top             =   2670
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 46"
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
         Height          =   255
         Index           =   46
         Left            =   -73650
         TabIndex        =   97
         Top             =   2400
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 45"
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
         Height          =   255
         Index           =   45
         Left            =   -73650
         TabIndex        =   96
         Top             =   2070
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 03"
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
         Height          =   255
         Index           =   3
         Left            =   -73650
         TabIndex        =   95
         Top             =   1740
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 02"
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
         Height          =   255
         Index           =   2
         Left            =   -73650
         TabIndex        =   94
         Top             =   1470
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 01"
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
         Height          =   255
         Index           =   1
         Left            =   -73650
         TabIndex        =   93
         Top             =   1200
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 00"
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
         Height          =   255
         Index           =   0
         Left            =   -73650
         TabIndex        =   92
         Top             =   870
         Width           =   645
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   91
         Tag             =   "v055"
         Top             =   870
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   -74880
         TabIndex        =   90
         Tag             =   "v056"
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   -74880
         TabIndex        =   89
         Tag             =   "v057"
         Top             =   1470
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   -74880
         TabIndex        =   88
         Tag             =   "v058"
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   45
         Left            =   -74880
         TabIndex        =   87
         Tag             =   "v059"
         Top             =   2070
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   46
         Left            =   -74880
         TabIndex        =   86
         Tag             =   "v060"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   47
         Left            =   -74880
         TabIndex        =   85
         Tag             =   "v061"
         Top             =   2670
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   48
         Left            =   -74880
         TabIndex        =   84
         Tag             =   "v062"
         Top             =   3000
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   49
         Left            =   -74880
         TabIndex        =   83
         Tag             =   "v063"
         Top             =   3270
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   54
         Left            =   -71610
         TabIndex        =   82
         Tag             =   "v064"
         Top             =   870
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   55
         Left            =   -71610
         TabIndex        =   81
         Tag             =   "v042"
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   56
         Left            =   -71610
         TabIndex        =   80
         Tag             =   "v043"
         Top             =   1410
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   57
         Left            =   -71610
         TabIndex        =   79
         Tag             =   "v044"
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   59
         Left            =   -71610
         TabIndex        =   78
         Tag             =   "v045"
         Top             =   3090
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   63
         Left            =   -71610
         TabIndex        =   77
         Tag             =   "v100"
         Top             =   2220
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   64
         Left            =   -71610
         TabIndex        =   76
         Tag             =   "v101"
         Top             =   3630
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   81
         Left            =   -74880
         TabIndex        =   75
         Tag             =   "v046"
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   82
         Left            =   -74880
         TabIndex        =   74
         Tag             =   "v047"
         Top             =   3870
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   83
         Left            =   -74880
         TabIndex        =   73
         Tag             =   "v048"
         Top             =   4140
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   84
         Left            =   -74880
         TabIndex        =   72
         Tag             =   "v050"
         Top             =   4470
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   85
         Left            =   -74880
         TabIndex        =   71
         Tag             =   "v051"
         Top             =   4740
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   86
         Left            =   -74880
         TabIndex        =   70
         Tag             =   "v052"
         Top             =   5070
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   87
         Left            =   -74880
         TabIndex        =   69
         Tag             =   "v053"
         Top             =   5340
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   88
         Left            =   -71610
         TabIndex        =   68
         Tag             =   "v054"
         Top             =   5430
         Width           =   1200
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 61"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   61
         Left            =   -70380
         TabIndex        =   67
         Top             =   1950
         Width           =   645
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   61
         Left            =   -71610
         TabIndex        =   66
         Top             =   1950
         Width           =   1200
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 65"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   65
         Left            =   -70380
         TabIndex        =   65
         Top             =   2490
         Width           =   645
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   65
         Left            =   -71610
         TabIndex        =   64
         Top             =   2490
         Width           =   1200
      End
      Begin VB.Label lblBEFVakXX 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -71610
         TabIndex        =   63
         Top             =   2760
         Width           =   1200
      End
      Begin VB.Label lblVakXXCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak XX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   -70380
         TabIndex        =   62
         Top             =   2760
         Width           =   645
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   62
         Left            =   -71610
         TabIndex        =   61
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 62"
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
         Height          =   255
         Index           =   62
         Left            =   -70380
         TabIndex        =   60
         Top             =   3360
         Width           =   645
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   66
         Left            =   -71610
         TabIndex        =   59
         Top             =   3900
         Width           =   1200
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 66"
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
         Height          =   255
         Index           =   66
         Left            =   -70380
         TabIndex        =   58
         Top             =   3900
         Width           =   645
      End
      Begin VB.Label lblBEFVakYY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -71610
         TabIndex        =   57
         Top             =   4170
         Width           =   1200
      End
      Begin VB.Label lblVakYYCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak YY"
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
         Height          =   255
         Index           =   0
         Left            =   -70380
         TabIndex        =   56
         Top             =   4170
         Width           =   645
      End
      Begin VB.Label lblBEF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "BEF "
         Height          =   195
         Index           =   0
         Left            =   -74220
         TabIndex        =   55
         Top             =   1260
         Width           =   525
      End
      Begin VB.Label lblBEF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "BEF "
         Height          =   195
         Index           =   1
         Left            =   -70920
         TabIndex        =   54
         Top             =   660
         Width           =   495
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   71
         Left            =   -71610
         TabIndex        =   53
         Top             =   4500
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   72
         Left            =   -71610
         TabIndex        =   52
         Top             =   4770
         Width           =   1200
      End
      Begin VB.Label lblBEFVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   91
         Left            =   -71610
         TabIndex        =   51
         Top             =   5100
         Width           =   1200
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 71"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   71
         Left            =   -70380
         TabIndex        =   50
         Top             =   4500
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 72"
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
         Height          =   255
         Index           =   72
         Left            =   -70380
         TabIndex        =   49
         Top             =   4770
         Width           =   645
      End
      Begin VB.Label lblVakCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vak 91"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   91
         Left            =   -70380
         TabIndex        =   48
         Top             =   5100
         Width           =   645
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -72990
         TabIndex        =   47
         Tag             =   "v055"
         Top             =   870
         Width           =   1200
      End
      Begin VB.Label lblBEF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "EUR "
         Height          =   195
         Index           =   2
         Left            =   -72360
         TabIndex        =   46
         Top             =   660
         Width           =   555
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   -72990
         TabIndex        =   45
         Tag             =   "v056"
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   -72990
         TabIndex        =   44
         Tag             =   "v057"
         Top             =   1470
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   -72990
         TabIndex        =   43
         Tag             =   "v058"
         Top             =   1740
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   45
         Left            =   -72990
         TabIndex        =   42
         Tag             =   "v059"
         Top             =   2070
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   46
         Left            =   -72990
         TabIndex        =   41
         Tag             =   "v060"
         Top             =   2400
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   47
         Left            =   -72990
         TabIndex        =   40
         Tag             =   "v061"
         Top             =   2670
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   48
         Left            =   -72990
         TabIndex        =   39
         Tag             =   "v062"
         Top             =   3000
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   49
         Left            =   -72990
         TabIndex        =   38
         Tag             =   "v063"
         Top             =   3270
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   81
         Left            =   -72990
         TabIndex        =   37
         Tag             =   "v046"
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   82
         Left            =   -72990
         TabIndex        =   36
         Tag             =   "v047"
         Top             =   3870
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   83
         Left            =   -72990
         TabIndex        =   35
         Tag             =   "v048"
         Top             =   4140
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   84
         Left            =   -72990
         TabIndex        =   34
         Tag             =   "v050"
         Top             =   4470
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   85
         Left            =   -72990
         TabIndex        =   33
         Tag             =   "v051"
         Top             =   4740
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   86
         Left            =   -72990
         TabIndex        =   32
         Tag             =   "v052"
         Top             =   5070
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   87
         Left            =   -72990
         TabIndex        =   31
         Tag             =   "v053"
         Top             =   5340
         Width           =   1200
      End
      Begin VB.Line Line1 
         X1              =   -71700
         X2              =   -71700
         Y1              =   870
         Y2              =   5580
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   54
         Left            =   -69720
         TabIndex        =   30
         Tag             =   "v064"
         Top             =   870
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   55
         Left            =   -69720
         TabIndex        =   29
         Tag             =   "v042"
         Top             =   1140
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   56
         Left            =   -69720
         TabIndex        =   28
         Tag             =   "v043"
         Top             =   1410
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   57
         Left            =   -69720
         TabIndex        =   27
         Tag             =   "v044"
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   61
         Left            =   -69720
         TabIndex        =   26
         Top             =   1950
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   63
         Left            =   -69720
         TabIndex        =   25
         Tag             =   "v100"
         Top             =   2220
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   65
         Left            =   -69720
         TabIndex        =   24
         Top             =   2490
         Width           =   1200
      End
      Begin VB.Label lblEURVakXX 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -69720
         TabIndex        =   23
         Top             =   2760
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   59
         Left            =   -69720
         TabIndex        =   22
         Tag             =   "v045"
         Top             =   3090
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   64
         Left            =   -69720
         TabIndex        =   21
         Tag             =   "v101"
         Top             =   3600
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   66
         Left            =   -69720
         TabIndex        =   20
         Top             =   3900
         Width           =   1200
      End
      Begin VB.Label lblEURVakYY 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -69720
         TabIndex        =   19
         Top             =   4170
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   71
         Left            =   -69720
         TabIndex        =   18
         Top             =   4500
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   72
         Left            =   -69720
         TabIndex        =   17
         Top             =   4770
         Width           =   1200
      End
      Begin VB.Label lblEURVak 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   91
         Left            =   -69780
         TabIndex        =   16
         Top             =   4140
         Width           =   1200
      End
      Begin VB.Label lblBEF 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "EUR "
         Height          =   195
         Index           =   3
         Left            =   -69090
         TabIndex        =   15
         Top             =   660
         Width           =   555
      End
   End
   Begin VB.CommandButton Afdrukken 
      Caption         =   "Afdruk"
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
      Left            =   6270
      TabIndex        =   11
      Top             =   30
      Width           =   1140
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   28
      Left            =   1770
      TabIndex        =   5
      Top             =   330
      Width           =   465
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   29
      Left            =   2310
      TabIndex        =   6
      Text            =   "31/12/1993"
      Top             =   330
      Width           =   1080
   End
   Begin VB.CommandButton Initialiseren 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Initialiseren"
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
      Left            =   6270
      TabIndex        =   3
      Top             =   750
      Width           =   1110
   End
   Begin VB.CommandButton Annuleren 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
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
      Left            =   6270
      TabIndex        =   0
      Top             =   360
      Width           =   1125
   End
   Begin MSComctlLib.TreeView tvwBtwAangiftes 
      Height          =   6645
      Left            =   60
      TabIndex        =   13
      Top             =   1140
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   11721
      _Version        =   393217
      Indentation     =   441
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   6
      Appearance      =   1
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
   Begin MSMAPI.MAPISession MPISessie 
      Left            =   120
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MPIBericht 
      Left            =   840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.CommandButton CmdEmailNBB 
      Caption         =   "vsoftBTW2007"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   240
      Picture         =   "frmBtwAangifte.frx":0B4F
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3000
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   4860
      TabIndex        =   10
      Top             =   600
      Width           =   1320
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   4860
      TabIndex        =   9
      Top             =   330
      Width           =   1320
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   3480
      TabIndex        =   8
      Top             =   600
      Width           =   1320
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   3480
      TabIndex        =   7
      Top             =   330
      Width           =   1320
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0000FF00&
      Caption         =   "Aankopen"
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
      Index           =   24
      Left            =   3600
      TabIndex        =   1
      Top             =   60
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Verkopen"
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
      Index           =   25
      Left            =   4980
      TabIndex        =   2
      Top             =   60
      Width           =   1200
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aktieve Periode tot :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   1770
      TabIndex        =   4
      Top             =   60
      Width           =   1605
   End
End
Attribute VB_Name = "frmBTWAangifte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Private mNode As node
Dim btwVakken As String

Dim rFlag As String

Dim vatDeclarationTemplate As String
Dim vatGridLineTemplate As String
Dim vatGridLinesList As String
Dim thisVatGridLine As String

Dim myVatNumber As String
Dim periodeFromRecord
Dim yearFromRecord
Dim myRegistrationName As String
Dim myStreetName As String
Dim myPostalZone As String
Dim myCityName As String
Dim myEmail As String
Dim periodeType As String
Dim gridNumber As String
Dim gridAmount As String

Private Sub Afdrukken_Click()

On Local Error Resume Next
PrintForm
If Err Then MsgBox Error

Annuleren.SetFocus

End Sub

Private Sub Annuleren_Click()

Unload Me

End Sub

Private Sub cbAanvraagBetaalformulieren_Click()

    Dim tmpXMLHier As String
    Dim tmpVakOud As String
    Dim tmpVakNieuw As String
    Dim ATOld As String
    Dim ATNew As String
    Dim ABOld As String
    Dim ABNew As String
    
    tmpXMLHier = RichTextBox1.text
    If Me.cbAanvraagTerugbetaling.Value = vbChecked Then
        ATOld = "YES"
    Else
        ATOld = "NO"
    End If
    ATNew = ATOld
    
    If Me.cbAanvraagBetaalformulieren.Value = vbChecked Then
        'nu wordt het YES, logisch zou vorige status NO zijn
        ABOld = "NO"
        ABNew = "YES"
    Else
        'en omgekeerd..
        ABOld = "YES"
        ABNew = "NO"
    End If
          
    tmpVakOud = "<ASK PAYMENT=" & Chr(34) & ABOld & Chr(34) & " RESTITUTION=" & Chr(34) & ATOld & Chr(34) & "/>"
    tmpVakNieuw = "<ASK PAYMENT=" & Chr(34) & ABNew & Chr(34) & " RESTITUTION=" & Chr(34) & ATNew & Chr(34) & "/>"
    tmpXMLHier = Replace(tmpXMLHier, tmpVakOud, tmpVakNieuw)
    
    RichTextBox1.text = tmpXMLHier

End Sub

Private Sub cbAanvraagTerugbetaling_Click()

    Dim tmpXMLHier As String
    Dim tmpVakOud As String
    Dim tmpVakNieuw As String
    Dim ATOld As String
    Dim ATNew As String
    Dim ABOld As String
    Dim ABNew As String
    
    tmpXMLHier = RichTextBox1.text
    If Me.cbAanvraagBetaalformulieren.Value = vbChecked Then
        ABOld = "YES"
    Else
        ABOld = "NO"
    End If
    ABNew = ABOld
    
    If Me.cbAanvraagTerugbetaling.Value = vbChecked Then
        'nu wordt het YES, logisch zou vorige status NO zijn
        ATOld = "NO"
        ATNew = "YES"
    Else
        'en omgekeerd..
        ATOld = "YES"
        ATNew = "NO"
    End If
          
    tmpVakOud = "<ASK PAYMENT=" & Chr(34) & ABOld & Chr(34) & " RESTITUTION=" & Chr(34) & ATOld & Chr(34) & "/>"
    tmpVakNieuw = "<ASK PAYMENT=" & Chr(34) & ABNew & Chr(34) & " RESTITUTION=" & Chr(34) & ATNew & Chr(34) & "/>"
    tmpXMLHier = Replace(tmpXMLHier, tmpVakOud, tmpVakNieuw)
    
    RichTextBox1.text = tmpXMLHier
    
End Sub


Private Sub cbIntervat2025_Click()
    
Dim vatFileName As String
vatFileName = myVatNumber + " " + Trim(Str(yearFromRecord)) + Dec((periodeFromRecord), "00") + ".xml"
    
XMLvraag2025:
Msg = "Btwaangifte doormailen voor afhandeling" & vbCrLf & vbCrLf & "Kies 'ja' voor doormailen (aanbevolen, ontvangstbewijs volgt), 'nee' indien U zelf het XML bestand afhandelt"
Ktrl = MsgBox(Msg, vbYesNoCancel + vbDefaultButton3 + vbQuestion)
Select Case Ktrl
    Case 2
        Msg = "Taak verlaten zonder enige verwerking.  Bent U zeker"
        Ktrl = MsgBox(Msg, vbYesNo + vbDefaultButton2 + vbQuestion)
        If Ktrl = 7 Then
            GoTo XMLvraag2025
        Else
            Exit Sub
        End If
        
    Case 6

        Screen.MousePointer = vbHourglass
        If Me.MPISessie.SessionID <> 0 Then
            MsgBox "Te controleren: MAPI sessie zou nog bezig zijn met iets anders.."
        Else
            On Local Error Resume Next
            Me.MPISessie.SignOn
            If Err Then
                SnelHelpPrint Err.Description, BL_LOGGING
            Else
                On Local Error GoTo 0
                Me.MPIBericht.SessionID = Me.MPISessie.SessionID
                SnelHelpPrint "E-mail sessie met succes opgestart. IDkode :" + Format(Me.MPISessie.SessionID), BL_LOGGING
            End If
        End If
        On Local Error GoTo MPIErrorXML2025
        'Compose new message
        Me.MPIBericht.Compose
        
        'Address message
        Me.MPIBericht.RecipDisplayName = "INTERVATBEHEER"
        Me.MPIBericht.RecipAddress = "SMTP:" & Me.tbmailbtw.text
        
        'Resolve recipient name
        Me.MPIBericht.AddressResolveUI = True '=dialogbox, false = error genereren
        
        'Create the message
        Me.MPIBericht.MsgSubject = "Verzoek tot BTW controle en aangifte"
        
        If Me.tbmailbtw.text = "info@rv.be" Then
            Me.MPIBericht.MsgNoteText = "Formaat:XML bestand" & vbCrLf & "In bijlage onze aangifte aangemaakt met marIntegraal versie " & MAR_VERSION & " voor controle en verzending.  Graag ontvangstbewijs binnen de 24 uur via mail of onze DNN postbus" & vbCrLf & vbCrLf & Now
        Else
            Me.MPIBericht.MsgNoteText = "Formaat:XML bestand" & vbCrLf & "In bijlage XML btw aangifte gegenereerd door ons boekhoudpakket.  Graag de aangifte door uw diensten na de gebruikelijke controles a.u.b." & vbCrLf & "Bezorgt U ons tevens nog ontvangstbevestiging ?" & vbCrLf & vbCrLf & "Dank bij voorbaat!" & vbCrLf & vbCrLf & Now
        End If
        
        Dim tmpXMLHier As String
        tmpXMLHier = Me.RichTextBox2.text
        
        Ktrl = ScrMaakTekstBestand(tmpXMLHier, vatFileName)
        Me.MPIBericht.AttachmentPathName = vatFileName
            
        'Send the message
        On Error Resume Next
        Me.MPIBericht.Send True
        Screen.MousePointer = vbNormal
        If Err = 32001 Then
            MsgBox "U koos om de mail af te breken.", vbInformation
            GoTo XMLvraag2025
        Else
            MsgBox "Zorg ervoor dat uw mailtoepassing effectief kan verzenden nu of straks.  U ontvangt later nog ontvangstbevestiging vanwege onze diensten", vbInformation
            Unload Me
            Exit Sub
        End If
           
    Case 7
        'niks
        'Zelf afhandelen (enkel nog afdruk afleveringnota's hierna
        Msg = "Klik het INTERVAT tabblad en bewaar het XML bestand (bvb. op uw bureaublad).  Start de INTERVAT webapplicatie en bezorg het XML bestand of breng uw cijfers manueel in in dezelfde toepassing" & vbCrLf & vbCrLf & "Voor hulp rond INTERVAT gelieve de website FOD te raadplegen."
        MsgBox Msg, vbInformation, "XML Btwaangifte zelf afhandelen via INTERVAT"
End Select
Exit Sub

MPIErrorXML2025:
MsgBox "error in mail"
    
End Sub

Private Sub cbPayment2025_Click()

    Dim tmpXmlToChange As String
    Dim paymentSearch As String
    Dim paymentChange As String
    Dim specialChar As String
    specialChar = """"
        
    tmpXmlToChange = Me.RichTextBox2.text
    If Me.cbPayment2025.Value = vbChecked Then
        'search for Payment="NO"
        paymentSearch = "Payment=" + specialChar + "NO" + specialChar
        paymentChange = "Payment=" + specialChar + "YES" + specialChar
        tmpXmlToChange = Replace(tmpXmlToChange, paymentSearch, paymentChange)
        Me.RichTextBox2.text = tmpXmlToChange
    Else
        'Payment="YES"
        paymentSearch = "Payment=" + specialChar + "YES" + specialChar
        paymentChange = "Payment=" + specialChar + "NO" + specialChar
        tmpXmlToChange = Replace(tmpXmlToChange, paymentSearch, paymentChange)
        Me.RichTextBox2.text = tmpXmlToChange
    End If


End Sub

Private Sub cbRestitution2025_Click()

    Dim tmpXmlToChange As String
    Dim restitutionSearch As String
    Dim restitutionChange As String
    Dim specialChar As String
    specialChar = """"
        
    tmpXmlToChange = Me.RichTextBox2.text
    If Me.cbRestitution2025.Value = vbChecked Then
        'search for Restitution="NO"
        restitutionSearch = "Restitution=" + specialChar + "NO" + specialChar
        restitutionChange = "Restitution=" + specialChar + "YES" + specialChar
        tmpXmlToChange = Replace(tmpXmlToChange, restitutionSearch, restitutionChange)
        Me.RichTextBox2.text = tmpXmlToChange
    Else
        'Restitution="YES"
        restitutionSearch = "Restitution=" + specialChar + "YES" + specialChar
        restitutionChange = "Restitution=" + specialChar + "NO" + specialChar
        tmpXmlToChange = Replace(tmpXmlToChange, restitutionSearch, restitutionChange)
        Me.RichTextBox2.text = tmpXmlToChange
    End If

End Sub

Private Sub cbVergrendel_Click()

    Me.RichTextBox1.Locked = Me.cbVergrendel.Value
    
End Sub

Private Sub cbXml2025_Click()
    
    Dim vatFileName As String
    
    vatFileName = myVatNumber + " " + Trim(Str(yearFromRecord)) + Dec((periodeFromRecord), "00") + ".xml"

    Mim.Teken.InitDir = ""
    Mim.Teken.filename = vatFileName
    Mim.Teken.ShowSave
    If Mim.Teken.filename = "" Then Exit Sub
    Dim tmpXMLHier As String
    tmpXMLHier = Me.RichTextBox2.text
    Ktrl = ScrMaakTekstBestand(tmpXMLHier, Mim.Teken.filename)

End Sub

Private Sub CmdEmailNBB_Click()

    Dim vatOut As String
    Dim vatTmp As String
    Dim BTWNummer As String
    Dim btwDatum As String
    Dim btwDatumFull As String
    Dim btwPeriode As String
    Dim btwWatIsDit As String
    Dim btwContactPersoon As String
    Dim btwContactTelefoon As String
    Dim btwContactMailAdres As String
    
    Select Case String99(READING, 301)
        Case "1", "2"
        Case Else
            MsgBox "Eerst Setup BTW instellen a.u.b.", vbInformation
            Exit Sub
    End Select
    
    If btwVakken = "" Then
        MsgBox "Eerst aankoop- en verkoopboeken afsluiten a.u.b. of selecteer eerst een bestaande aangifte", vbInformation
        Exit Sub
    End If
        
    BTWNummer = BtwKontrole(String99(READING, 51), True)
    btwDatum = Format(TekstInfo(29), "yymmdd")
    btwDatumFull = Format(TekstInfo(29), "yyyymmdd")
    btwWatIsDit = "777"
    btwPeriode = ""
    btwContactPersoon = String99(READING, 52)
    btwContactTelefoon = String99(READING, 49)
    btwContactMailAdres = String99(READING, 50) '= mail!
        
    Dim btwOnderwerp As String
    Dim vatNewSubject As String
        
    btwOnderwerp = "$marNT$BTWAGF$" & BTWNummer & "$" & btwPeriode
    vatTmp = "UNB*UNOA:3*" & BTWNummer & "*VAT-ADMIN*" & btwDatum & ":1246*" & BTWNummer & _
                "00001*" & btwWatIsDit & "*VATDEC-EUR***1.0" & _
                "'UNH*1*RDRMES:D:96A:UN:VAT001" & _
                "'BGM*937*124980407*2" & _
                "'DTM*137:" & btwDatumFull & ":102" & _
                "'DTM*320:" & btwPeriode & ":608" & _
                "'NAD*DT*" & BTWNummer & ":52:129" & _
                "'CTA*IC*:" & btwContactPersoon & _
                "'COM*" & btwContactTelefoon & ":TE" & _
                "'COM*" & btwContactMailAdres & ":ML" & _
                "'IDE*1*VATDEC" & _
                "'NAD*DT*" & BTWNummer & ":52:129" & _
                "'FTX*AAI***0000:EUR" & _
                btwVakken & _
                "'UNT*32*1" & _
                "'UNZ*1*" & BTWNummer & "00001'" & _
                "'RVS:" & String99(READING, 301) & "*" & String99(READING, 302) & "*" & String99(READING, 303) & "'"
                
    
                                
    MsgBox "Dit venster wordt hierna gesloten en de gestructureerde aangifte verzonden via de ingebouwde Internet Explorer.  Wacht in de 'Explorer' op bevestiging van verzending a.u.b !", vbExclamation

Err = 0
On Error Resume Next
            
            
If Err Then MsgBox Error: Exit Sub
Unload Me

End Sub


Private Sub Command1_Click()

    Mim.Teken.InitDir = ""
    Mim.Teken.filename = "btw.xml"
    Mim.Teken.ShowSave
    If Mim.Teken.filename = "" Then Exit Sub
    Dim tmpXMLHier As String
    tmpXMLHier = Me.RichTextBox1.text
    Ktrl = ScrMaakTekstBestand(tmpXMLHier, Mim.Teken.filename)
    
End Sub

Private Sub cbIntervat2008_Click()

XMLvraag:
Msg = "Btwaangifte doormailen voor afhandeling" & vbCrLf & vbCrLf & "Kies 'ja' voor doormailen (aanbevolen, ontvangstbewijs volgt), 'nee' indien U zelf het XML bestand afhandelt"
Ktrl = MsgBox(Msg, vbYesNoCancel + vbDefaultButton3 + vbQuestion)
Select Case Ktrl
    Case 2
        Msg = "Taak verlaten zonder enige verwerking.  Bent U zeker"
        Ktrl = MsgBox(Msg, vbYesNo + vbDefaultButton2 + vbQuestion)
        If Ktrl = 7 Then
            GoTo XMLvraag
        Else
            Exit Sub
        End If
        
    Case 6

        Screen.MousePointer = vbHourglass
        If Me.MPISessie.SessionID <> 0 Then
            MsgBox "Te controleren: MAPI sessie zou nog bezig zijn met iets anders.."
        Else
            On Local Error Resume Next
            Me.MPISessie.SignOn
            If Err Then
                SnelHelpPrint Err.Description, BL_LOGGING
            Else
                On Local Error GoTo 0
                Me.MPIBericht.SessionID = Me.MPISessie.SessionID
                SnelHelpPrint "E-mail sessie met succes opgestart. IDkode :" + Format(Me.MPISessie.SessionID), BL_LOGGING
            End If
        End If
        On Local Error GoTo MPIErrorXML
        'Compose new message
        Me.MPIBericht.Compose
        
        'Address message
        Me.MPIBericht.RecipDisplayName = "INTERVATBEHEER"
        Me.MPIBericht.RecipAddress = "SMTP:" & Me.tbmailbtw.text
        
        'Resolve recipient name
        Me.MPIBericht.AddressResolveUI = True '=dialogbox, false = error genereren
        
        'Create the message
        Me.MPIBericht.MsgSubject = "Verzoek tot BTW controle en aangifte"
        
        If Me.tbmailbtw.text = "info@rv.be" Then
            Me.MPIBericht.MsgNoteText = "Formaat:XML bestand" & vbCrLf & "In bijlage onze aangifte aangemaakt met marIntegraal versie " & MAR_VERSION & " voor controle en verzending.  Graag ontvangstbewijs binnen de 24 uur via mail of onze DNN postbus" & vbCrLf & vbCrLf & Now
        Else
            Me.MPIBericht.MsgNoteText = "Formaat:XML bestand" & vbCrLf & "In bijlage XML btw aangifte gegenereerd door ons boekhoudpakket.  Graag de aangifte door uw diensten na de gebruikelijke controles a.u.b." & vbCrLf & "Bezorgt U ons tevens nog ontvangstbevestiging ?" & vbCrLf & vbCrLf & "Dank bij voorbaat!" & vbCrLf & vbCrLf & Now
        End If
        
        Dim tmpXMLHier As String
        tmpXMLHier = Me.RichTextBox1.text
        Ktrl = ScrMaakTekstBestand(tmpXMLHier, "btwaangifte.xml")
        '

        Me.MPIBericht.AttachmentPathName = "btwaangifte.xml"
            
        'Send the message
        On Error Resume Next
        Me.MPIBericht.Send True
        Screen.MousePointer = vbNormal
        If Err = 32001 Then
            MsgBox "U koos om de mail af te breken.", vbInformation
            GoTo XMLvraag
        Else
            MsgBox "Zorg ervoor dat uw mailtoepassing effectief kan verzenden nu of straks.  U ontvangt later nog ontvangstbevestiging vanwege onze diensten", vbInformation
            Unload Me
            Exit Sub
        End If
           
    Case 7
        'niks
        'Zelf afhandelen (enkel nog afdruk afleveringnota's hierna
        Msg = "Klik het INTERVAT tabblad en bewaar het XML bestand (bvb. op uw bureaublad).  Start de INTERVAT webapplicatie en bezorg het XML bestand of breng uw cijfers manueel in in dezelfde toepassing" & vbCrLf & vbCrLf & "Voor hulp rond INTERVAT gelieve de website FOD te raadplegen."
        MsgBox Msg, vbInformation, "XML Btwaangifte zelf afhandelen via INTERVAT"
End Select
Exit Sub

MPIErrorXML:
MsgBox "error in mail"

End Sub



Function XmlGenerate() As Boolean

XmlGenerate = False
Ktrl = ScrLeesBestandAlleTekst(vatDeclarationTemplate, PROGRAM_LOCATION + "xml-templates\vat\be-vatdeclare.xml")
If Ktrl = 0 Then
    MsgBox "Onverwachte situatie", vbCritical
End If
vatGridLineTemplate = ""
vatGridLinesList = ""

vatDeclarationTemplate = Replace(vatDeclarationTemplate, "{vatNumberDeclarant}", myVatNumber)
vatDeclarationTemplate = Replace(vatDeclarationTemplate, "{nameDeclarant}", myRegistrationName)
vatDeclarationTemplate = Replace(vatDeclarationTemplate, "{streetDeclarant}", myStreetName)
vatDeclarationTemplate = Replace(vatDeclarationTemplate, "{postalCodeDeclarant}", myPostalZone)
vatDeclarationTemplate = Replace(vatDeclarationTemplate, "{cityDeclarant}", myCityName)
vatDeclarationTemplate = Replace(vatDeclarationTemplate, "{emailDeclarant}", myEmail)

yearFromRecord = Trim(Str(Val(vBibTekst(TABLE_VARIOUS, "i002"))))
Select Case String99(READING, 301)
    Case "2"    '
        periodeType = "Quarter"
        periodeFromRecord = Trim(Str(Int(Val(vBibTekst(TABLE_VARIOUS, "i001")) / 3)))
                    
    Case "1"
        periodeType = "Month"
        periodeFromRecord = Trim(Str(Val(vBibTekst(TABLE_VARIOUS, "i001"))))
        
    Case Else
        MsgBox "Setup BTW instellen a.u.b.", vbCritical
        Me.SSTab1.TabVisible(2) = False
        Me.SSTab1.TabVisible(3) = False
        Me.cbIntervat2008.Enabled = False
        Me.cbIntervat2025.Enabled = False
End Select

vatDeclarationTemplate = Replace(vatDeclarationTemplate, "{periodeType}", periodeType)
vatDeclarationTemplate = Replace(vatDeclarationTemplate, "{period}", periodeFromRecord)
vatDeclarationTemplate = Replace(vatDeclarationTemplate, "{fiscalYear}", yearFromRecord)
vatDeclarationTemplate = Replace(vatDeclarationTemplate, "{referenceDeclarant}", myVatNumber + "-" + Trim(Str(yearFromRecord)) + Trim(Dec((periodeFromRecord), "00")))
'Me.RichTextBox2.Text = vatDeclarationTemplate
XmlGenerate = True

End Function



Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

lblInfoTekst = "Vak 54 t.e.m. Vak 72 geeft de som weer resulterende uit afdruk van aankoop- en verkoopboeken (werkwijze 1985 - 2003)." & _
                "rVak 54 t.e.m. rVak 72 geeft som van de cijfers teruggevonden op de boekhoudrekeningen voor Vak54 tot Vak72 !" & vbCrLf & vbCrLf & _
                "AANDACHT:" & vbCrLf & "De cijfers in rVak54 tot rVak72 worden aangewend voor GESTRUCTUREERDE AANGIFTE !!"

Set mNode = tvwBtwAangiftes.nodes.Add() ' Create the first node.
mNode.text = "Btw aangiftes"

myVatNumber = Trim(String99(READING, 51))
myRegistrationName = Trim(String99(READING, 46))
myRegistrationName = CheckforAmp(myRegistrationName)

myStreetName = Trim(String99(READING, 47))
myStreetName = CheckforAmp(myStreetName)

If Mid(String99(READING, 48), 5, 1) = " " Then
    myPostalZone = Mid(String99(READING, 48), 1, 4)
    myCityName = Trim(Mid(String99(READING, 48), 5))
Else
    MsgBox "Controleer setup en parameters voor postcode en plaatsnaam." & vbCrLf & vbCrLf & "Postcode uit 4 cijfers gevold door een spatie.", vbExclamation
End If
myCompanyId = supplierCountryCode + supplierRegistrationId
myEmail = Trim(String99(READING, 50))


Dim dokumentTot As Long

bGetOrGreater TABLE_VARIOUS, 1, vSet("17", 20)
If Ktrl Then
Else
    Do
        RecordToVeld TABLE_VARIOUS
        dokumentTot = 0
        For COUNT_TO = 92 To 98 Step 2
            dokumentTot = dokumentTot + Val(vBibTekst(TABLE_VARIOUS, "#v" + Format(COUNT_TO, "000") + " #"))
        Next
        If dokumentTot Then
            Set mNode = tvwBtwAangiftes.nodes.Add(1, tvwChild)
            mNode.text = vBibTekst(TABLE_VARIOUS, "#v090 #") + " " + vBibTekst(TABLE_VARIOUS, "#v091 #")
        End If
        bNext TABLE_VARIOUS
        If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "17" Then
            Exit Do
        End If
    Loop
End If
SnelHelpPrint Now & ":In verdere ontwikkeling voor BEF/EUR/E-mail", BL_LOGGING
VulDeVelden BJPERDAT.Boekjaar.text, Format(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00")
'BJPERDAT.Boekjaar.Text + Format(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00")
End Sub

Private Sub Form_Unload(Cancel As Integer)

bClose TABLE_VARIOUS

End Sub

Private Sub Initialiseren_Click()
ReDim pTrec(7) As Integer
Dim Nrr             As Long
Dim PeriodeSleutel  As String * 20
Dim DummySleutel    As String * 5
Dim getal As Long

Msg = "BTW Aangifte periode initializeren." + vbCrLf + "Bent U zeker ?"
Ktrl = MsgBox(Msg$, vbQuestion + vbYesNo + vbDefaultButton2, "Alle boeken opnieuw uitdrukken !?!")
If Ktrl = vbYes Then
Else
    Exit Sub
End If

Ktrl = 0
SnelHelpPrint "Kontrole hogere periodes...", BL_LOGGING
PeriodeMax = BJPERDAT.PeriodeBoekjaar.ListCount + 1
Do While PeriodeMax > BJPERDAT.PeriodeBoekjaar.ListIndex + 1
    PeriodeSleutel = "17" + BJPERDAT.Boekjaar.text + Format(PeriodeMax, "00")
    bGet TABLE_VARIOUS, 1, PeriodeSleutel
    If Ktrl Then
        TLB_RECORD(TABLE_VARIOUS) = ""
        vBib TABLE_VARIOUS, (BJPERDAT.Boekjaar.text), "v090"
        vBib TABLE_VARIOUS, Format(PeriodeMax, "00"), "v091"
        vBib TABLE_VARIOUS, "17" + vBibTekst(TABLE_VARIOUS, "#v090 #") + vBibTekst(TABLE_VARIOUS, "#v091 #"), "v005"
        bInsert TABLE_VARIOUS, 1
    Else
        RecordToVeld TABLE_VARIOUS
        getal = 0
        
        For T = 92 To 99
            getal = getal + Val(vBibTekst(TABLE_VARIOUS, "#v" + Format(T, "000") + " #"))
        Next
        If getal Then
            getal = PeriodeMax
            PeriodeMax = 0
            Exit Do
        End If
    End If
    PeriodeMax = PeriodeMax - 1
Loop

jump:
If getal Then
    MsgBox "Periode " + Format(getal, "00") + " reeds afgesloten..."
    Initialiseren.Enabled = False
    CmdEmailNBB.Enabled = False
    Exit Sub
Else
    PeriodeSleutel = "17" + BJPERDAT.Boekjaar.text + Format(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00")
    bGet TABLE_VARIOUS, 1, PeriodeSleutel
    If Ktrl Then
        TLB_RECORD(TABLE_VARIOUS) = ""
        vBib TABLE_VARIOUS, (BJPERDAT.Boekjaar.text), "v090"
        vBib TABLE_VARIOUS, Format(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00"), "v091"
        vBib TABLE_VARIOUS, "17" + vBibTekst(TABLE_VARIOUS, "#v090 #") + vBibTekst(TABLE_VARIOUS, "#v091 #"), "v005"
        bInsert TABLE_VARIOUS, 1
        GoTo jump
    Else
        RecordToVeld TABLE_VARIOUS
    End If
End If

pTrec(1) = 2
pTrec(3) = 4
pTrec(5) = 12
pTrec(7) = 14
For T = 1 To 8 Step 2
    If Val(vBibTekst(TABLE_VARIOUS, "#v" + Format(T + 91, "000") + " #")) = 0 Then
    Else
        Nrr = Val(vBibTekst(TABLE_VARIOUS, "#v" + Format(T + 91, "000") + " #")) - 1
        SS99 Format(Nrr, "00000"), pTrec(T)
    End If
Next
TLB_RECORD(TABLE_VARIOUS) = ""
vBib TABLE_VARIOUS, (BJPERDAT.Boekjaar.text), "v090"
vBib TABLE_VARIOUS, Format(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00"), "v091"
vBib TABLE_VARIOUS, "17" + vBibTekst(TABLE_VARIOUS, "#v090 #") + vBibTekst(TABLE_VARIOUS, "#v091 #"), "v005"
bUpdate TABLE_VARIOUS, 1
Initialiseren.Enabled = False
CmdEmailNBB.Enabled = False
VulDeVelden BJPERDAT.Boekjaar.text, Format(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00")
Annuleren.SetFocus

End Sub

Private Sub VulDeVelden(strJaar As String, strPeriode As String)

btwVakken = ""
bGet TABLE_VARIOUS, 1, vSet("17" & strJaar & strPeriode, 20)
If Ktrl Then
    Beep
    Me.SSTab1.TabEnabled(2) = False
    Exit Sub
Else
    RecordToVeld TABLE_VARIOUS
End If

TekstInfo(28).text = Format(BJPERDAT.PeriodeBoekjaar.ListIndex + 1, "00")
TekstInfo(29).text = Right(BJPERDAT.PeriodeBoekjaar.text, 10)

Label2(0).Caption = Format(Val(vBibTekst(TABLE_VARIOUS, "#v092 #")), "00000") + " - " + Format(Val(vBibTekst(TABLE_VARIOUS, "#v093 #")), "00000")
Label2(2).Caption = Format(Val(vBibTekst(TABLE_VARIOUS, "#v096 #")), "00000") + " - " + Format(Val(vBibTekst(TABLE_VARIOUS, "#v097 #")), "00000")
Label2(1).Caption = Format(Val(vBibTekst(TABLE_VARIOUS, "#v094 #")), "00000") + " - " + Format(Val(vBibTekst(TABLE_VARIOUS, "#v095 #")), "00000")
Label2(3).Caption = Format(Val(vBibTekst(TABLE_VARIOUS, "#v098 #")), "00000") + " - " + Format(Val(vBibTekst(TABLE_VARIOUS, "#v099 #")), "00000")

If vBibTekst(TABLE_VARIOUS, "i001") = "" Then
    MsgBox "Aan- en verkoopboeken voor Intervat nog af te drukken.  Intervat tabblad blijft uitgeschakeld.  Intervat enkel mogelijk met boeken uitgedrukt via versie 6.6.900 of hoger.  TIP: initialiseer fiche en druk de boeken opnieuw uit met versie 900 of hoger", vbInformation
    Me.SSTab1.TabVisible(2) = False
    Me.cbIntervat2008.Enabled = False
    Me.cbIntervat2025.Enabled = False
Else
    Me.SSTab1.TabVisible(2) = True
    Me.cbIntervat2008.Enabled = True
    Me.cbIntervat2025.Enabled = True
End If

Ktrl = XmlGenerate()
If Ktrl = 0 Then
    MsgBox "Fout bij afhandelen xml bestand", vbCritical
End If

Dim tmpXMLHier As String
Dim tmpVak As String
Dim tmpVakHier As String

INTERVAT2008:
Me.RichTextBox1.LoadFile App.path & "\XMLbtwAangifte.rtf"
tmpXMLHier = RichTextBox1.text

tmpVak = "<VATNUMBER>9999999999</VATNUMBER>"
tmpVakHier = String99(READING, 51)
If Len(tmpVakHier) <> 10 Then
    MsgBox "btw notatie in setup dient 10 cijfers te bevatten startend met 0 of 1 (vb. 0440058217). Verbeter setup vooraleer verder te gaan", vbExclamation
    Me.SSTab1.TabVisible(2) = False
    Me.cbIntervat2008.Enabled = False
    Exit Sub
Else
    tmpVakHier = "0" & Mid(tmpVakHier, 1, 3) & Mid(tmpVakHier, 5, 3) & Mid(tmpVakHier, 9, 3)
    tmpXMLHier = Replace(tmpXMLHier, tmpVak, "<VATNUMBER>" & tmpVakHier & "</VATNUMBER>")
    Me.SSTab1.TabVisible(2) = True
    Me.cbIntervat2008.Enabled = True
End If

tmpVak = "<NAME>Contactpersoon</NAME>"
tmpVakHier = String99(READING, 52)
tmpVakHier = CheckforAmp(tmpVakHier)
tmpXMLHier = Replace(tmpXMLHier, tmpVak, "<NAME>" & tmpVakHier & "</NAME>")

tmpVak = "<ADDRESS>StraatContact</ADDRESS>"
tmpVakHier = String99(READING, 47)
tmpVakHier = CheckforAmp(tmpVakHier)
tmpXMLHier = Replace(tmpXMLHier, tmpVak, "<ADDRESS>" & tmpVakHier & "</ADDRESS>")

tmpVak = "<POSTCODE>0000</POSTCODE>"
tmpVakHier = Left(String99(READING, 48), 4)
tmpXMLHier = Replace(tmpXMLHier, tmpVak, "<POSTCODE>" & tmpVakHier & "</POSTCODE>")

tmpVak = "<CITY>Plaatscontact</CITY>"
tmpVakHier = Trim(Mid(String99(READING, 48), 5))
tmpXMLHier = Replace(tmpXMLHier, tmpVak, "<CITY>" & tmpVakHier & "</CITY>")

tmpVak = "<SENDINGREFERENCE>99999999900000</SENDINGREFERENCE>"
tmpVakHier = String99(READING, 51)
If Len(tmpVakHier) <> 10 Then
    MsgBox "btw notatie in setup dient 10 cijfers te bevatten startend met 0 of 1 (vb. 0440058217). Verbeter setup vooraleer verder te gaan", vbExclamation
    Me.SSTab1.TabVisible(2) = False
    Me.cbIntervat2008.Enabled = False
    Exit Sub
Else
    tmpVakHier = Mid(tmpVakHier, 1, 3) & Mid(tmpVakHier, 5, 3) & Mid(tmpVakHier, 9, 3) & "00000"
    tmpXMLHier = Replace(tmpXMLHier, tmpVak, "<SENDINGREFERENCE>" & tmpVakHier & "</SENDINGREFERENCE>")
    Me.SSTab1.TabVisible(2) = True
    Me.cbIntervat2008.Enabled = True
End If

tmpVak = "<VATNUMBER>0000000000</VATNUMBER>"
tmpVakHier = String99(READING, 51)
If Len(tmpVakHier) <> 10 Then
    MsgBox "btw notatie in setup dient 10 cijfers te bevatten startend met 0 of 1 (vb. 0440058217). Verbeter setup vooraleer verder te gaan", vbExclamation
    Me.SSTab1.TabVisible(2) = False
    Me.cbIntervat2008.Enabled = False
    Exit Sub
Else
    tmpVakHier = "0" & Mid(tmpVakHier, 1, 3) & Mid(tmpVakHier, 5, 3) & Mid(tmpVakHier, 9, 3)
    tmpXMLHier = Replace(tmpXMLHier, tmpVak, "<VATNUMBER>" & tmpVakHier & "</VATNUMBER>")
    Me.SSTab1.TabVisible(2) = True
    Me.cbIntervat2008.Enabled = True
End If
                
tmpVak = "<NAME>NaamBedrijf</NAME>"
tmpVakHier = String99(READING, 46)
tmpVakHier = CheckforAmp(tmpVakHier)
tmpXMLHier = Replace(tmpXMLHier, tmpVak, "<NAME>" & tmpVakHier & "</NAME>")
                
tmpVak = "<ADDRESS>StraatBedrijf</ADDRESS>"
tmpVakHier = String99(READING, 47)
tmpVakHier = CheckforAmp(tmpVakHier)
tmpXMLHier = Replace(tmpXMLHier, tmpVak, "<ADDRESS>" & tmpVakHier & "</ADDRESS>")

tmpVak = "<POSTCODE>9999</POSTCODE>"
tmpVakHier = Left(String99(READING, 48), 4)
tmpXMLHier = Replace(tmpXMLHier, tmpVak, "<POSTCODE>" & tmpVakHier & "</POSTCODE>")

tmpVak = "<CITY>PlaatsBedrijf</CITY>"
tmpVakHier = Trim(Mid(String99(READING, 48), 5))
tmpXMLHier = Replace(tmpXMLHier, tmpVak, "<CITY>" & tmpVakHier & "</CITY>")

Select Case String99(READING, 301)
    Case "2"
        tmpXMLHier = Replace(tmpXMLHier, "<QUARTERORMONTH>0</QUARTERORMONTH>", "<QUARTER>" & Trim(Str(Int(Val(vBibTekst(TABLE_VARIOUS, "i001")) / 3))) & "</QUARTER>")
                    
    Case "1"
        tmpXMLHier = Replace(tmpXMLHier, "<QUARTERORMONTH>0</QUARTERORMONTH>", "<MONTH>" & Trim(Str(Val(vBibTekst(TABLE_VARIOUS, "i001")))) & "</MONTH>")
    Case Else
        MsgBox "Setup BTW instellen a.u.b.", vbCritical
        Me.SSTab1.TabVisible(2) = False
        Me.SSTab1.TabVisible(3) = False
        Me.cbIntervat2008.Enabled = False
        Me.cbIntervat2025.Enabled = False
End Select
tmpXMLHier = Replace(tmpXMLHier, "<YEAR>1985</YEAR>", "<YEAR>" & Trim(Str(Val(vBibTekst(TABLE_VARIOUS, "i002")))) & "</YEAR>")


On Error Resume Next

lblBEFVak(61).Caption = Format(0, "#,##0")
lblBEFVak(62).Caption = Format(0, "#,##0")
lblEURVak(61).Caption = Format(0, "#,##0.00")
lblEvak(61).Caption = Format(0, "#,##0.00")
lblEURVak(62).Caption = Format(0, "#,##0.00")
lblEvak(62).Caption = Format(0, "#,##0.00")

Ktrl = ScrLeesBestandAlleTekst(vatGridLineTemplate, PROGRAM_LOCATION + "xml-templates\vat\be-vatgridline.xml")
If Ktrl = 0 Then
    MsgBox "Onverwachte situatie", vbCritical
End If

If vBibTekst(TABLE_VARIOUS, "#vEUR #") = "EUR" Then
    For T = 0 To lblBEFVak.UBound
        If lblBEFVak(T).Tag <> "" Then
            'Debug.Print lblBEFVak(T).Tag
            lblBEFVak(T).Caption = Format(Val(vBibTekst(TABLE_VARIOUS, "#" + lblBEFVak(T).Tag + " #")) * EURO, "#,##0")
            lblEURVak(T).Caption = Format(Val(vBibTekst(TABLE_VARIOUS, "#" + lblBEFVak(T).Tag + " #")), "#,##0.00")
            lblEvak(T).Caption = lblEURVak(T).Caption
            If CCur(lblEvak(T).Caption) = 0 Then
            ''Debug.Print lblEvak(T).Caption, CCur(lblEvak(T).Caption)
            Else
                btwVakken = btwVakken & "'SCD*3*" & Format(T, "00") & _
                    "'ARR**" & Trim$(Str$(Val(vBibTekst(TABLE_VARIOUS, "#" + lblBEFVak(T).Tag + " #")) * 100))
                tmpVak = "<D" & Trim(Str(T)) & ">0</D" & Trim(Str(T)) & ">"
                tmpVakHier = "<D" & Trim(Str(T)) & ">" & Trim$(Str$(Val(vBibTekst(TABLE_VARIOUS, "#" + lblBEFVak(T).Tag + " #")) * 100)) & "</D" & Trim(Str(T)) & ">"
                tmpXMLHier = Replace(tmpXMLHier, tmpVak, tmpVakHier)
                                
                'amount
                thisVatGridLine = vatGridLineTemplate
                gridAmount = Trim(Dec(Val(vBibTekst(TABLE_VARIOUS, "#" + lblBEFVak(T).Tag + " #")), MASK_EUR))
                thisVatGridLine = Replace(thisVatGridLine, "{amount}", gridAmount)
                'gridNumber
                gridNumber = Trim(Str(T))
                thisVatGridLine = Replace(thisVatGridLine, "{gridnumber}", gridNumber)
                vatGridLinesList = vatGridLinesList + thisVatGridLine + vbCrLf
            End If
        End If
    Next
Else
    For T = 0 To lblBEFVak.UBound
        If lblBEFVak(T).Tag <> "" Then
            lblBEFVak(T).Caption = Format(Val(vBibTekst(TABLE_VARIOUS, "#" + lblBEFVak(T).Tag + " #")), "#,##0")
            lblEURVak(T).Caption = Format(Val(vBibTekst(TABLE_VARIOUS, "#" + lblBEFVak(T).Tag + " #")) / EURO, "#,##0.00")
            lblEvak(T).Caption = lblEURVak(T).Caption
        End If
    Next
End If

On Local Error GoTo 0
Me.lblEURVakXX(0).Caption = Format(CCur(Me.lblEURVak(54)) + CCur(Me.lblEURVak(55)) + CCur(Me.lblEURVak(56)) + CCur(Me.lblEURVak(57)) + CCur(Me.lblEURVak(61)) + CCur(Me.lblEURVak(63)), "#,##0.00")
Me.lblEURVakXX(1).Caption = Me.lblEURVakXX(0).Caption
Me.lblEURVakYY(0).Caption = Format(CCur(Me.lblEURVak(59)) + CCur(Me.lblEURVak(62)) + CCur(Me.lblEURVak(64)), "#,##0.00")
Me.lblEURVakYY(1).Caption = Me.lblEURVakYY(0).Caption

'nu bepalen 71 of 72
'Me.cbRestitution2025.Value = Checked
If CCur(lblEURVakXX(1)) - CCur(lblEURVakYY(1)) < 0 Then
    lblEvak(71).Caption = Format(0, "#,##0.00")
    lblEvak(72).Caption = Format(Abs(CCur(lblEURVakXX(1)) - CCur(lblEURVakYY(1))), "#,##0.00")
    tmpVak = "<D71>0</D71>"
    tmpVakHier = "<D72>" & Trim(Str(CCur(lblEvak(72).Caption) * 100)) & "</D72>"
    tmpXMLHier = Replace(tmpXMLHier, tmpVak, tmpVakHier)
    
    'amount
    thisVatGridLine = vatGridLineTemplate
    gridAmount = Trim(Str(CCur(lblEvak(72).Caption)))
    thisVatGridLine = Replace(thisVatGridLine, "{amount}", gridAmount)
    'gridNumber
    gridNumber = "72"
    thisVatGridLine = Replace(thisVatGridLine, "{gridnumber}", gridNumber)
    vatGridLinesList = vatGridLinesList + thisVatGridLine
    'Me.cbRestitution2025.Value = Unchecked
Else
    lblEvak(72).Caption = Format(0, "#,##0.00")
    lblEvak(71).Caption = Format(Abs(CCur(lblEURVakXX(1)) - CCur(lblEURVakYY(1))), "#,##0.00")
    
    tmpVak = "<D71>0</D71>"
    tmpVakHier = "<D71>" & Trim(Str(CCur(lblEvak(71).Caption) * 100)) & "</D71>"
    tmpXMLHier = Replace(tmpXMLHier, tmpVak, tmpVakHier)
    
    'amount
    thisVatGridLine = vatGridLineTemplate
    gridAmount = Trim(Str(CCur(lblEvak(71).Caption)))
    thisVatGridLine = Replace(thisVatGridLine, "{amount}", gridAmount)
    'gridNumber
    gridNumber = "71"
    thisVatGridLine = Replace(thisVatGridLine, "{gridnumber}", gridNumber)
    vatGridLinesList = vatGridLinesList + thisVatGridLine
    
End If
RichTextBox1.text = tmpXMLHier

'<Vsoft>vatDeclareGridLinesList</Vsoft>
vatDeclarationTemplate = Replace(vatDeclarationTemplate, "<Vsoft>vatDeclareGridLinesList</Vsoft>", vatGridLinesList)
Me.RichTextBox2.text = vatDeclarationTemplate

If Left(tmpVakHier, 5) = "</D72>" Then
    '72 is teruggave dus aanvraag teruggave automatisch aanklikken
    Me.cbAanvraagTerugbetaling.Value = vbChecked
    Me.cbRestitution2025.Value = vbChecked
End If

End Sub



Private Sub SSTab1_Click(PreviousTab As Integer)

    'Select Case SSTab1.Tab
    '    Case 2
    '        Select Case String99(READING, 301)
    '            Case "2"
    '                Me.RichTextBox1.LoadFile App.path & "\QbtwAangifte.rtf"
    '
    '            Case "1"
    '                Me.RichTextBox1.LoadFile App.path & "\MbtwAangifte.rtf"
    '            Case Else
    '                MsgBox "Setup BTW instellen a.u.b.", vbCritical
    '        End Select
    '        RichTextBox1.Text = Replace(RichTextBox1.Text, "<D3>0.00</D3>", "<D3>1.11</D3>")
    'End Select

End Sub


Private Sub tvwBtwAangiftes_NodeClick(ByVal node As MSComctlLib.node)
   
    If node.Index > 1 Then
        'MsgBox "Index = " & Node.Index & " Text:" & Node.Text
        Initialiseren.Enabled = False
        CmdEmailNBB.Enabled = False
        VulDeVelden Left(node.text, 4), Right(node.text, 2)
    End If
    
End Sub
