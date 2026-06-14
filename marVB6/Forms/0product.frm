VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmProduktFiche 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProduktFiche"
   ClientHeight    =   5940
   ClientLeft      =   3135
   ClientTop       =   3300
   ClientWidth     =   9240
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
   ScaleHeight     =   5940
   ScaleWidth      =   9240
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab v 
      Height          =   5775
      Left            =   0
      TabIndex        =   58
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10186
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Default"
      TabPicture(0)   =   "0product.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(24)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(19)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Line1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblCijfers(6)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(22)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LblCijfers(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LblCijfers(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LblCijfers(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "LblCijfers(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LblCijfers(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LblCijfers(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(21)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Line1(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(25)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(29)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(28)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(26)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(27)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(14)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(32)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(13)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(15)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(16)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(17)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(11)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(10)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(9)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(7)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label1(8)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label1(23)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Label1(3)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label1(4)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label1(5)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Label1(6)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label1(1)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Label1(2)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label1(31)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label1(20)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Label1(18)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label1(0)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "lbJournaal"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Label1(30)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Label1(33)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtEindeReeks"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "CmdDommeKey"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "TxtInfo(0)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "TxtInfo(1)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "TxtInfo(3)"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "TxtInfo(2)"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "TxtInfo(5)"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "TxtInfo(7)"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "TxtInfo(6)"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "TxtInfo(4)"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "TxtInfo(10)"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "TxtInfo(9)"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "TxtInfo(8)"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "TxtInfo(12)"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "TxtInfo(11)"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "TxtInfo(13)"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "TxtInfo(15)"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "TxtInfo(14)"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "TxtInfo(16)"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "TxtInfo(18)"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "TxtInfo(17)"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "TxtInfo(19)"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "TxtInfo(20)"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "TxtInfo(21)"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "TxtInfo(22)"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "TxtInfo(23)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "cbCategorie"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "chkFilter(0)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "chkFilter(1)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "chkFilter(2)"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "chkFilter(3)"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).Control(75)=   "cmdSwitch"
      Tab(0).Control(75).Enabled=   0   'False
      Tab(0).Control(76)=   "txtLink"
      Tab(0).Control(76).Enabled=   0   'False
      Tab(0).Control(77)=   "cmdTonen"
      Tab(0).Control(77).Enabled=   0   'False
      Tab(0).Control(78)=   "cmdSchoon"
      Tab(0).Control(78).Enabled=   0   'False
      Tab(0).Control(79)=   "CmdVerwijderFiche"
      Tab(0).Control(79).Enabled=   0   'False
      Tab(0).Control(80)=   "cmdLager"
      Tab(0).Control(80).Enabled=   0   'False
      Tab(0).Control(81)=   "CmdPrinterAfdruk"
      Tab(0).Control(81).Enabled=   0   'False
      Tab(0).Control(82)=   "CmdBewaren"
      Tab(0).Control(82).Enabled=   0   'False
      Tab(0).Control(83)=   "CmdHoger"
      Tab(0).Control(83).Enabled=   0   'False
      Tab(0).Control(84)=   "CmdSluiten"
      Tab(0).Control(84).Enabled=   0   'False
      Tab(0).Control(85)=   "Alfa"
      Tab(0).Control(85).Enabled=   0   'False
      Tab(0).Control(86)=   "cbMerk"
      Tab(0).Control(86).Enabled=   0   'False
      Tab(0).Control(87)=   "txtMilieu"
      Tab(0).Control(87).Enabled=   0   'False
      Tab(0).Control(88)=   "cmdRBAcontrole"
      Tab(0).Control(88).Enabled=   0   'False
      Tab(0).ControlCount=   89
      TabCaption(1)   =   "SQL Query"
      TabPicture(1)   =   "0product.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblRecordCount"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "msfSQL"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txtSQL"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdSQL"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmdKopij"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "eCommerce FTP"
      TabPicture(2)   =   "0product.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Journaal"
      TabPicture(3)   =   "0product.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "msfJournaal"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.CommandButton cmdRBAcontrole 
         Caption         =   "&Controle"
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
         Left            =   6480
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   4200
         Width           =   735
      End
      Begin VB.TextBox txtMilieu 
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
         Left            =   2520
         TabIndex        =   49
         Top             =   4200
         Width           =   3855
      End
      Begin VB.CommandButton cmdKopij 
         Caption         =   "XML &Kopie"
         Height          =   690
         Left            =   -66000
         Picture         =   "0product.frx":0070
         Style           =   1  'Graphical
         TabIndex        =   88
         ToolTipText     =   "XML RecordSet maken of kopiëren naar klassiek plakbord"
         Top             =   1560
         Width           =   1200
      End
      Begin VB.CommandButton cmdSQL 
         Caption         =   "SQL &SELECT"
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
         Left            =   -74760
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtSQL 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1395
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   84
         Text            =   "0product.frx":06DA
         Top             =   120
         Width           =   10185
      End
      Begin VB.ComboBox cbMerk 
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
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   4950
         Width           =   3015
      End
      Begin VB.CommandButton Alfa 
         Caption         =   "SQL &Zoeken"
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
         Left            =   7320
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   1590
         Width           =   1605
      End
      Begin VB.CommandButton CmdSluiten 
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
         Left            =   7320
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   2430
         Width           =   1605
      End
      Begin VB.CommandButton CmdHoger 
         Caption         =   "&Hoger"
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
         Left            =   7320
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   810
         Width           =   1605
      End
      Begin VB.CommandButton CmdBewaren 
         Caption         =   "&Bewaren"
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
         Left            =   7320
         TabIndex        =   50
         Top             =   2070
         Width           =   1605
      End
      Begin VB.CommandButton CmdPrinterAfdruk 
         Caption         =   "&Printer Afdruk"
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
         Left            =   7320
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   2790
         Width           =   1605
      End
      Begin VB.CommandButton cmdLager 
         Caption         =   "&Lager"
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
         Left            =   7320
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   450
         Width           =   1605
      End
      Begin VB.CommandButton CmdVerwijderFiche 
         Caption         =   "&Verwijderen"
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
         Left            =   7320
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1605
      End
      Begin VB.CommandButton cmdSchoon 
         Caption         =   "Schoon"
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
         Left            =   7320
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   30
         Width           =   1605
      End
      Begin VB.CommandButton cmdTonen 
         Caption         =   "&Tonen"
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
         Left            =   6480
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   4560
         Width           =   735
      End
      Begin VB.TextBox txtLink 
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
         Left            =   480
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   4560
         Width           =   5895
      End
      Begin VB.CommandButton cmdSwitch 
         Caption         =   "Ingave in EUR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4680
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   120
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Verkoop Inclusief"
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
         Index           =   3
         Left            =   7320
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   4200
         Width           =   1590
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Aankoop Inclusief"
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
         Index           =   2
         Left            =   7320
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   3960
         Width           =   1590
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Winstberekening"
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
         Index           =   1
         Left            =   7320
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1500
      End
      Begin VB.CheckBox chkFilter 
         Caption         =   "Per Verpakking"
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
         Index           =   0
         Left            =   7320
         TabIndex        =   59
         Top             =   3510
         Width           =   1500
      End
      Begin VB.ComboBox cbCategorie 
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
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   4950
         Width           =   2895
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   23
         Left            =   2760
         TabIndex        =   47
         Tag             =   "v103"
         ToolTipText     =   "v103"
         Top             =   3870
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   22
         Left            =   630
         TabIndex        =   45
         Tag             =   "v125"
         ToolTipText     =   "v125"
         Top             =   3870
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   21
         Left            =   4290
         TabIndex        =   43
         Tag             =   "v104"
         ToolTipText     =   "v104"
         Top             =   3600
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   20
         Left            =   1170
         TabIndex        =   41
         Tag             =   "v124;&L"
         ToolTipText     =   "v124"
         Top             =   3600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   19
         Left            =   5970
         TabIndex        =   39
         Tag             =   "e123|v123"
         ToolTipText     =   "e123|v123"
         Top             =   3240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   17
         Left            =   5970
         TabIndex        =   35
         Tag             =   "e121|v121"
         ToolTipText     =   "e121|v121"
         Top             =   2640
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   18
         Left            =   5970
         TabIndex        =   37
         Tag             =   "e122|v122"
         ToolTipText     =   "e122|v122"
         Top             =   2940
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   16
         Left            =   3750
         TabIndex        =   33
         Tag             =   "v114"
         ToolTipText     =   "v114"
         Top             =   3240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   14
         Left            =   3750
         TabIndex        =   29
         Tag             =   "v119"
         ToolTipText     =   "v119"
         Top             =   2640
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   15
         Left            =   3750
         TabIndex        =   31
         Tag             =   "v120"
         ToolTipText     =   "v120"
         Top             =   2940
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   13
         Left            =   1230
         TabIndex        =   27
         Tag             =   "v118;&R3"
         ToolTipText     =   "v118"
         Top             =   3240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   11
         Left            =   1230
         TabIndex        =   23
         Tag             =   "v116;&R6"
         ToolTipText     =   "v116"
         Top             =   2640
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   12
         Left            =   1230
         TabIndex        =   25
         Tag             =   "v117;&R7"
         ToolTipText     =   "v117"
         Top             =   2940
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   8
         Left            =   990
         TabIndex        =   17
         Tag             =   "e112|v112"
         ToolTipText     =   "e112|v112"
         Top             =   2250
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   9
         Left            =   3390
         TabIndex        =   19
         Tag             =   "e113|v113"
         ToolTipText     =   "e113|v113"
         Top             =   2250
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   10
         Left            =   6210
         TabIndex        =   21
         Tag             =   "v115"
         ToolTipText     =   "v115"
         Top             =   2250
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   4
         Left            =   5070
         TabIndex        =   9
         Tag             =   "v108;022"
         ToolTipText     =   "v108"
         Top             =   1650
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   6
         Left            =   3990
         TabIndex        =   13
         Tag             =   "v110"
         ToolTipText     =   "v110"
         Top             =   1950
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   7
         Left            =   5850
         TabIndex        =   15
         Tag             =   "v111;002"
         ToolTipText     =   "v111"
         Top             =   1950
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   5
         Left            =   750
         TabIndex        =   11
         Tag             =   "v109"
         ToolTipText     =   "v109"
         Top             =   1950
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   2
         Left            =   930
         TabIndex        =   5
         Tag             =   "v106;004"
         ToolTipText     =   "v106"
         Top             =   1650
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   3
         Left            =   3480
         TabIndex        =   7
         Tag             =   "v107"
         ToolTipText     =   "v107"
         Top             =   1650
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   1
         Left            =   2280
         TabIndex        =   3
         Tag             =   "v105"
         ToolTipText     =   "v105"
         Top             =   1290
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   40
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TxtInfo 
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Tag             =   "v102"
         ToolTipText     =   "v102"
         Top             =   90
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   13
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfSQL 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   86
         Top             =   2280
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   4895
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfJournaal 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   90
         Top             =   120
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   8493
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.CommandButton CmdDommeKey 
         Caption         =   "DommeKey"
         Default         =   -1  'True
         Height          =   195
         Left            =   6600
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   4680
         Width           =   375
      End
      Begin MSMask.MaskEdBox txtEindeReeks 
         Height          =   255
         Left            =   6120
         TabIndex        =   93
         Tag             =   "v107"
         ToolTipText     =   "v107"
         Top             =   840
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Einde reeks korting%"
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
         Index           =   33
         Left            =   5520
         TabIndex        =   92
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Y. Recupel;Bebat;Auvibel enz."
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
         Index           =   30
         Left            =   30
         TabIndex        =   48
         Top             =   4200
         Width           =   2475
      End
      Begin VB.Label lbJournaal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3600
         TabIndex        =   89
         ToolTipText     =   "Gejournaliseerd in actief boekjaar"
         Top             =   120
         Width           =   1065
      End
      Begin VB.Label lblRecordCount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   -67440
         TabIndex        =   87
         Top             =   1560
         Width           =   1185
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&X. Kode (EAN)"
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
         Index           =   0
         Left            =   60
         TabIndex        =   0
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Gemiddelde Aankoopwaarde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   18
         Left            =   60
         TabIndex        =   83
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EUR excl. Btw"
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
         Index           =   20
         Left            =   2640
         TabIndex        =   82
         Top             =   630
         Width           =   1245
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&A. Nederlandse omschrijving"
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
         Index           =   31
         Left            =   30
         TabIndex        =   2
         Top             =   1290
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&C. Verpakking"
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
         Index           =   2
         Left            =   2370
         TabIndex        =   6
         Top             =   1650
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&B. Maatstaf"
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
         Index           =   1
         Left            =   30
         TabIndex        =   4
         Top             =   1650
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&E. Plaats"
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
         Index           =   6
         Left            =   30
         TabIndex        =   10
         Top             =   1950
         Width           =   690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&G. Btw Kode"
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
         Index           =   5
         Left            =   4890
         TabIndex        =   14
         Top             =   1950
         Width           =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&F. Winst %"
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
         Index           =   4
         Left            =   3135
         TabIndex        =   12
         Top             =   1950
         Width           =   810
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&D. Soort"
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
         Index           =   3
         Left            =   4410
         TabIndex        =   8
         Top             =   1650
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&J. Minimum Stock"
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
         Index           =   23
         Left            =   4890
         TabIndex        =   20
         Top             =   2250
         Width           =   1305
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&I. Aankoop"
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
         Index           =   8
         Left            =   2490
         TabIndex        =   18
         Top             =   2250
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&H. Verkoop"
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
         Index           =   7
         Left            =   30
         TabIndex        =   16
         Top             =   2250
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&K. Aankooprek."
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
         Index           =   9
         Left            =   30
         TabIndex        =   22
         Top             =   2640
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&L. Verkooprek."
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
         Index           =   10
         Left            =   30
         TabIndex        =   24
         Top             =   2940
         Width           =   1155
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&M. Voorraadrek."
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
         Index           =   11
         Left            =   15
         TabIndex        =   26
         Top             =   3240
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&O. Eenheden"
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
         Index           =   17
         Left            =   2655
         TabIndex        =   30
         Top             =   2940
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&N. Eenheden"
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
         Index           =   16
         Left            =   2670
         TabIndex        =   28
         Top             =   2640
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&P. Eenheden"
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
         Index           =   15
         Left            =   2670
         TabIndex        =   32
         Top             =   3240
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Q. Bedrag"
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
         Index           =   13
         Left            =   5190
         TabIndex        =   34
         Top             =   2640
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&R. Bedrag"
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
         Index           =   32
         Left            =   5190
         TabIndex        =   36
         Top             =   2940
         Width           =   780
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&S. Bedrag"
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
         Index           =   14
         Left            =   5205
         TabIndex        =   38
         Top             =   3240
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&T. Leverancier"
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
         Index           =   27
         Left            =   30
         TabIndex        =   40
         Top             =   3600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&T 2e kodenr."
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
         Index           =   26
         Left            =   3270
         TabIndex        =   42
         Top             =   3600
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&U. Vlag"
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
         Index           =   28
         Left            =   30
         TabIndex        =   44
         Top             =   3870
         Width           =   585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Link"
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
         Index           =   29
         Left            =   0
         TabIndex        =   51
         Top             =   4590
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&V. Goederenkode"
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
         Left            =   1200
         TabIndex        =   46
         Top             =   3870
         Width           =   1320
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   30
         X2              =   7170
         Y1              =   1230
         Y2              =   1230
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EUR incl. Btw"
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
         Left            =   2640
         TabIndex        =   81
         Top             =   930
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Input&Fillter Bedragen"
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
         Index           =   21
         Left            =   7290
         TabIndex        =   80
         Top             =   3180
         Width           =   1515
      End
      Begin VB.Label LblCijfers 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   315
         Index           =   0
         Left            =   5760
         TabIndex        =   79
         ToolTipText     =   "Actuele Stock"
         Top             =   150
         Width           =   1395
      End
      Begin VB.Label LblCijfers 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   315
         Index           =   1
         Left            =   60
         TabIndex        =   78
         Top             =   870
         Width           =   1215
      End
      Begin VB.Label LblCijfers 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Index           =   2
         Left            =   5400
         TabIndex        =   77
         Top             =   5520
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label LblCijfers 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Index           =   3
         Left            =   6720
         TabIndex        =   76
         Top             =   5520
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.Label LblCijfers 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Index           =   4
         Left            =   3960
         TabIndex        =   75
         Top             =   930
         Width           =   1245
      End
      Begin VB.Label LblCijfers 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   285
         Index           =   5
         Left            =   3960
         TabIndex        =   74
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Actuele Aankoopwaarde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   22
         Left            =   1320
         TabIndex        =   73
         Top             =   390
         Width           =   1275
      End
      Begin VB.Label LblCijfers 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         Height          =   315
         Index           =   6
         Left            =   1320
         TabIndex        =   72
         Top             =   870
         Width           =   1275
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   30
         X2              =   7170
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Categorie"
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
         Index           =   19
         Left            =   0
         TabIndex        =   54
         Top             =   4950
         Width           =   1215
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Merk"
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
         Left            =   7320
         TabIndex        =   57
         Top             =   4590
         Width           =   1215
      End
   End
   Begin VB.Menu Opties 
      Caption         =   "&Opties"
      Begin VB.Menu LijstRap 
         Caption         =   "&Lijstrapportage"
      End
      Begin VB.Menu VerwijderenMogelijk 
         Caption         =   "&Verwijderen mogelijk maken"
      End
      Begin VB.Menu Groepen 
         Caption         =   "&Groepen"
      End
   End
End
Attribute VB_Name = "frmProduktFiche"
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

Dim FicheIsNieuw As Boolean
Dim IsGewijzigd As Boolean
Dim iTabIndex           As Integer

Dim TempoMilieu As String

Dim rsGroepenHier As ADODB.Recordset
Dim rsSQLQuery As ADODB.Recordset

Dim rsJourQuery As ADODB.Recordset

Sub RecordVullen()
Dim TempoTel As Integer
Dim vBibDef As String

Err = 0
TempoTel = 0
On Local Error Resume Next
Do
    If TxtInfo(TempoTel).Tag = "" Then
    Else
        If InStr(TxtInfo(TempoTel).Tag, "|") <> 0 Then
            If InStr(cmdSwitch.Caption, "BEF") Then
                vBibDef = Left(TxtInfo(TempoTel).Tag, 4)
                vBib TABLE_PRODUCTS, Dec(Val(TxtInfo(TempoTel).text) / EURO, MASK_EURX), vBibDef
                vBibDef = Right(TxtInfo(TempoTel).Tag, 4)
            Else
                vBibDef = Right(TxtInfo(TempoTel).Tag, 4)
                vBib TABLE_PRODUCTS, Dec(Val(TxtInfo(TempoTel).text) * EURO, MASK_EURX), vBibDef
                vBibDef = Left(TxtInfo(TempoTel).Tag, 4)
            End If
        ElseIf InStr(TxtInfo(TempoTel).Tag, ";") <> 0 Then
            vBibDef = Left(TxtInfo(TempoTel).Tag, InStr(TxtInfo(TempoTel).Tag, ";") - 1)
        Else
            vBibDef = TxtInfo(TempoTel).Tag
        End If
        vBib TABLE_PRODUCTS, TxtInfo(TempoTel).text, vBibDef
    End If
    TempoTel = TempoTel + 1
    SnelHelpPrint TxtInfo(TempoTel).Tag, BL_LOGGING
    If Err Then Exit Do
Loop
vBib TABLE_PRODUCTS, Me.txtLink.text, "v002"
vBib TABLE_PRODUCTS, Me.txtMilieu.text, "v261"
vBib TABLE_PRODUCTS, Me.txtEindeReeks.text, "v300"
If Me.cbCategorie.ListIndex >= 0 Then
    vBib TABLE_PRODUCTS, Me.cbCategorie.text, "v221"
Else
    vBib TABLE_PRODUCTS, " ", "v221"
End If
If Me.cbMerk.ListIndex >= 0 Then
    vBib TABLE_PRODUCTS, Me.cbMerk.text, "v001"
Else
    vBib TABLE_PRODUCTS, " ", "v001"
End If
   
End Sub

Sub Schoon()
Dim TempoTel As Integer

TempoTel = 24
Do While TempoTel > -1
    On Error Resume Next
    TxtInfo(TempoTel) = ""
    TempoTel = TempoTel - 1
Loop
Me.txtLink.text = ""
Me.txtMilieu.text = ""
frmProduktFiche.Caption = "ProduktFiche :"

'v106
TxtInfo(2) = fmarBoxText("004", "2", "0")
'v107
TxtInfo(3) = Dec$(1, "#####.00")
'v108
TxtInfo(4) = fmarBoxText("022", "2", "N")
'v111
TxtInfo(7) = fmarBoxText("002", "2", String99(READING, 183))
'v116
TxtInfo(11) = String99(READING, 77)
'v117
TxtInfo(12) = String99(READING, 78)
'v118
TxtInfo(13) = String99(READING, 79)

VernieuwStockWaarde
VernieuwTicketPrijs

On Error Resume Next
'TxtInfo(0).SetFocus
Me.cbCategorie.ListIndex = -1
Me.cbMerk.ListIndex = -1

End Sub

Sub VernieuwStockWaarde()
Dim TotaalBeweging As Double
Dim TotaalAantal As Single

'Dim Stock As cmGemiddelde
'Set Stock = New cmGemiddelde
'Stock.aantalElementen = 2
'Stock.AddItem 2366
'Stock.AddItem 7889
'If Stock.Volledig Then
'    MsgBox Stock.GeMiddelde
'End If
'Set Stock = Nothing

TotaalAantal = Val(TxtInfo(14)) + Val(TxtInfo(15)) + Val(TxtInfo(16))
TotaalBeweging = Val(TxtInfo(17)) + Val(TxtInfo(18)) + Val(TxtInfo(19))

LblCijfers(0) = Format(Val(TxtInfo(14)) + Val(TxtInfo(16)) - Val(TxtInfo(15)), "0.000")
If TotaalAantal = 0 Then
    LblCijfers(1) = ""
Else
    LblCijfers(1) = Format((TotaalBeweging / TotaalAantal) * CDbl(LblCijfers(0)), "#,##0.00")
End If
LblCijfers(6) = Format(Val(TxtInfo(9)) * CDbl(LblCijfers(0)), "#,##0.00")

End Sub

Sub VernieuwTicketPrijs()

    Dim BedragBefExcl As Currency
    Dim BedragBefIncl As Currency
    Dim BedragBefBtw  As Currency

    Dim BedragEurExcl As Currency
    Dim BedragEurIncl As Currency
    Dim bedragEurBtw  As Currency

'maskerEur
If InStr(cmdSwitch.Caption, "EUR") Then
    BedragEurExcl = Val(TxtInfo(3).text) * Val(TxtInfo(8).text)
    BedragBefExcl = BedragEurExcl * EURO
    
    bedragEurBtw = BedragEurExcl * Val(Mid(TxtInfo(7), InStr(TxtInfo(7), ":") + 1)) / 100
    BedragBefBtw = BedragBefExcl * Val(Mid(TxtInfo(7), InStr(TxtInfo(7), ":") + 1)) / 100
    
    LblCijfers(5) = Format(BedragEurExcl, "#,##0.00")
    LblCijfers(3) = Format(BedragBefExcl, "#,##0.00")
    LblCijfers(4) = Format(BedragEurExcl + bedragEurBtw, "#,##0.00")
    LblCijfers(2) = Format(BedragBefExcl + BedragBefBtw, "#,##0")
Else
    BedragBefExcl = Val(TxtInfo(3).text) * Val(TxtInfo(8).text)
    BedragEurExcl = BedragBefExcl / EURO
    
    BedragBefBtw = CDbl(BedragBefExcl * Val(Mid(TxtInfo(7), InStr(TxtInfo(7), ":") + 1)) / 100)
    bedragEurBtw = CDbl(BedragEurExcl * Val(Mid(TxtInfo(7), InStr(TxtInfo(7), ":") + 1)) / 100)
    
    LblCijfers(3) = Format(BedragBefExcl, "#,##0.00")
    LblCijfers(5) = Format(BedragEurExcl, "#,##0.00")
    LblCijfers(2) = Format(BedragBefExcl + BedragBefBtw, "#,##0")
    LblCijfers(4) = Format(BedragEurExcl + bedragEurBtw, "#,##0.00")
End If

End Sub

Private Sub Alfa_Click()

SharedFl = TABLE_PRODUCTS

aIndex = 1
GridText = TxtInfo(1).text
SqlSearch.Show 1
If Ktrl Then
Else
    VensterVullen
End If

End Sub

Private Sub chkFilter_Click(Index As Integer)

If Index = 1 Then
    If chkFilter(1) = 1 Then
        TxtInfo(8).Enabled = False
    Else
        TxtInfo(8).Enabled = True
    End If
End If

End Sub

Private Sub CmdDommeKey_Click()
Dim CheckBool As Boolean

VolgendVeld:
    Err = 0
    On Local Error Resume Next
    CheckBool = TxtInfo(iTabIndex + 1).TabStop
    If Err Then
        cmdBewaren.SetFocus
    Else
        If CheckBool = True Then
            TxtInfo(iTabIndex + 1).SetFocus
        Else
            GoTo VolgendVeld
        End If
    End If
    
End Sub

Private Sub CmdHoger_Click()

bNext TABLE_PRODUCTS
If Ktrl Then
    bLast TABLE_PRODUCTS, 0
    If Ktrl Then
        Beep
        MsgBox "Er zijn nog geen gegevens..."
        Exit Sub
    End If
End If
VensterVullen
        
End Sub

Private Sub cmdKopij_Click()

Msg = "Kies 'Ja' voor kopij als XML bestand" & vbCr & _
    "Kies 'Nee' voor kopij naar het klassieke plakbord"

KtrlBox = MsgBox(Msg, vbQuestion + vbYesNoCancel + vbDefaultButton3)
If KtrlBox = vbCancel Then
    Exit Sub
ElseIf KtrlBox = vbNo Then
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText msfSQL.Clip
    Exit Sub
Else
    On Error GoTo CancelError
    Mim.Teken.fileName = ""
    Mim.Teken.CancelError = True
    Mim.Teken.Filter = "Alle bestanden (*.xml)|*.xml"
    Mim.Teken.ShowSave
    Dim BestandHier As String
    BestandHier = Mim.Teken.fileName
    If Not Dir(BestandHier) = "" Then
        Kill BestandHier
    End If
    rsSQLQuery.Save Mim.Teken.fileName, adPersistXML
End If

CancelError:
Screen.MousePointer = vbNormal

End Sub

Private Sub CmdLager_Click()

bPrev TABLE_PRODUCTS
If Ktrl Then
    bFirst TABLE_PRODUCTS, 0
    If Ktrl Then
        Beep
        MsgBox "Er zijn nog geen gegevens..."
        Exit Sub
    End If
End If
VensterVullen

End Sub


Private Sub cmdRBAcontrole_Click()

Dim tempo() As String
Dim COUNT_TO As Integer
Dim TelOK As Integer

TempoMilieu = txtMilieu.text
TelOK = -1

tempo = Split(txtMilieu.text, ";")
For COUNT_TO = 0 To UBound(tempo)
    bGet TABLE_PRODUCTS, 0, tempo(COUNT_TO)
    If Ktrl = 0 Then
        MsgBox Trim(tempo(COUNT_TO)) & " aanwezig", vbInformation
        TelOK = TelOK + 1
    Else
        MsgBox Trim(tempo(COUNT_TO)) & " NIET aanwezig", vbExclamation
    End If
Next
If TelOK = UBound(tempo) Then

Else
    TempoMilieu = ""
    MsgBox "Probeer opnieuw"
    
End If

End Sub

Private Sub cmdSluiten_Click()

Unload Me

End Sub


Private Sub cmdBewaren_Click()
Dim TempoFL As Integer
Dim TeSwitchen As Boolean

TeSwitchen = False
If bhEuro Then
    If InStr(cmdSwitch.Caption, "BEF") Then
        SnelHelpPrint "BEF switch voor EURO boekhouding.  Cijfers werden omgewerkt vooraleer weg te schrijven", BL_LOGGING
        cmdSwitch_Click
        TeSwitchen = True
    End If
ElseIf InStr(cmdSwitch.Caption, "EUR") Then
    SnelHelpPrint "EUR switch voor BEF boekhouding.  Cijfers worden omgewerkt vooraleer weg te schrijven", BL_LOGGING
    cmdSwitch_Click
    TeSwitchen = True
End If

If TempoMilieu = "" Then
Else
    txtMilieu.text = TempoMilieu
    TempoMilieu = ""
End If
Msg = "Gegevens bestaande '" + bstNaam(TABLE_PRODUCTS) + "'-fiche wijzigen.  Bent U zeker ?"
KtrlBox = MsgBox(Msg, 292)
If KtrlBox = vbYes Then
    RecordVullen
    If TeSwitchen = True Then
        cmdSwitch_Click
    End If
    
    If FicheIsNieuw Then
        bInsert TABLE_PRODUCTS, 0
        If frmProduktFiche.BorderStyle Then
            cmdSchoon_Click
        Else
            Unload Me
        End If
    Else
        bUpdate TABLE_PRODUCTS, 0
    End If
Else
    If TeSwitchen = True Then
        cmdSwitch_Click
    End If
End If


End Sub


Private Sub VensterVullen()
Dim TempoTel As Integer
Dim vBibDef As String

FicheIsNieuw = False
RecordToVeld TABLE_PRODUCTS
frmProduktFiche.Caption = "ProduktFiche : " + RTrim$(vBibTekst(TABLE_PRODUCTS, "#v105 #"))

'If rsPFStat.State = adStateOpen Then
'        rsPFStat.Close
'    End If
'    Msg = "SELECT * FROM ProductenTB2 WHERE Polis = '" & TxtInfo(0).Text & "'"
'    rsTB2.Open Msg, adntDB, adOpenDynamic, adLockOptimistic, adCmdText
'    Me.cbPolisHistoriek.Clear
'    If rsTB2.RecordCount = 0 Then
'        SSTab2.TabCaption(1) = "&TB2= (0)"
'        Me.rtbIndentTB2.Text = ""
'    Else
'        SSTab2.TabCaption(1) = "&TB2= (" & Trim$(Str$(rsTB2.RecordCount)) & ")"
'        Do While Not rsTB2.EOF
'            strTB2 = rsTB2("memoTB2").Value
'            Me.cbPolisHistoriek.AddItem "Aanvangsdatum " & rsTB2("DatumAanvang").Value
 '           rsTB2.MoveNext
'        Loop
'    End If
'    'rsAllerlei.Close
'End If

TempoTel = 0
On Local Error Resume Next
For TempoTel = 0 To TxtInfo.UBound
    If TxtInfo(TempoTel).Tag = "" Then
    Else
        If InStr(TxtInfo(TempoTel).Tag, "|") <> 0 Then
            If bhEuro Then
                'boekhoudcijfers in EUR
                vBibDef = Left(TxtInfo(TempoTel).Tag, 4)
                If InStr(cmdSwitch.Caption, "EUR") Then
                    TxtInfo(TempoTel).text = Dec(Val(vBibTekst(TABLE_PRODUCTS, "#" + vBibDef + " #")), "#######0.0000")
                Else
                    TxtInfo(TempoTel).text = Dec(Val(vBibTekst(TABLE_PRODUCTS, "#" + vBibDef + " #")) * EURO, "#######0.0000")
                End If
            Else
                'boekhoudcijfers in BEF
                vBibDef = Right(TxtInfo(TempoTel).Tag, 4)
                If InStr(cmdSwitch.Caption, "BEF") Then
                    TxtInfo(TempoTel).text = Dec(Val(vBibTekst(TABLE_PRODUCTS, "#" + vBibDef + " #")), "#######0.0000")
                Else
                    TxtInfo(TempoTel).text = Dec(Val(vBibTekst(TABLE_PRODUCTS, "#" + vBibDef + " #")) / EURO, "#######0.0000")
                End If
            End If
        ElseIf InStr(TxtInfo(TempoTel).Tag, ";") <> 0 Then
            vBibDef = Left(TxtInfo(TempoTel).Tag, InStr(TxtInfo(TempoTel).Tag, ";") - 1)
            If InStr(TxtInfo(TempoTel).Tag, "&") Then
                TxtInfo(TempoTel).text = vBibTekst(TABLE_PRODUCTS, "#" + vBibDef + " #")
            Else
                TxtInfo(TempoTel).text = fmarBoxText(Mid(TxtInfo(TempoTel).Tag, InStr(TxtInfo(TempoTel).Tag, ";") + 1), "2", vBibTekst(TABLE_PRODUCTS, "#" + vBibDef + " #"))
            End If
        Else
            vBibDef = TxtInfo(TempoTel).Tag
            TxtInfo(TempoTel).text = vBibTekst(TABLE_PRODUCTS, "#" + vBibDef + " #")
        End If
    End If
    'If TempoTel = 16 Then Stop
Next
Me.txtLink.text = vBibTekst(TABLE_PRODUCTS, "#v002 #")
Me.txtEindeReeks.text = Str(Val(vBibTekst(TABLE_PRODUCTS, "#v300 #")))
Me.txtMilieu.text = vBibTekst(TABLE_PRODUCTS, "#v261 #")
Dim tmpCategorie As String
Dim tmpMerk As String
Dim COUNT_TOHier As Integer

tmpCategorie = Trim$(vBibTekst(TABLE_PRODUCTS, "#v221 #"))
tmpMerk = Trim$(vBibTekst(TABLE_PRODUCTS, "#v001 #"))
Me.cbMerk.ListIndex = -1
Me.cbCategorie.ListIndex = -1

If tmpCategorie = "" Then
    'niks
Else
    For COUNT_TOHier = 0 To Me.cbCategorie.ListCount - 1
        If tmpCategorie = Me.cbCategorie.List(COUNT_TOHier) Then
            Me.cbCategorie.ListIndex = COUNT_TOHier
        End If
    Next
End If

If tmpMerk = "" Then
    'niks
Else
    For COUNT_TOHier = 0 To Me.cbMerk.ListCount - 1
        If tmpMerk = Me.cbMerk.List(COUNT_TOHier) Then
            Me.cbMerk.ListIndex = COUNT_TOHier
        End If
    Next
End If

VernieuwTicketPrijs
VernieuwStockWaarde

End Sub


Private Sub cmdSQL_Click()

Dim retCode As Boolean

retCode = adoRECORDset

End Sub

Function adoRECORDset() As Boolean

    Screen.MousePointer = vbHourglass
    
    On Error Resume Next
    msfSQL.Recordset = Nothing
    
    rsSQLQuery.Close
    On Error Resume Next
    rsSQLQuery.Open txtSQL.text, adntDB
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        msfSQL.Refresh
    Else
        Set msfSQL.DataSource = rsSQLQuery
    End If
    lblRecordCount = rsSQLQuery.RecordCount
    Screen.MousePointer = vbNormal

End Function


Private Sub cmdSwitch_Click()
Dim TempoTel As Integer

If cmdSwitch.Caption = "Ingave in EUR" Then
    cmdSwitch.Caption = "Ingave in BEF"
    'de nodige schermopmaak voor BEF
    On Local Error Resume Next
    For TempoTel = 0 To TxtInfo.UBound
        If InStr(TxtInfo(TempoTel).Tag, "|") <> 0 Then
            TxtInfo(TempoTel).text = Dec(Val(TxtInfo(TempoTel).text) * EURO, "#######0.0000")
        End If
    Next
Else
    cmdSwitch.Caption = "Ingave in EUR"
    'de nodige schermopmaak voor EUR
    On Local Error Resume Next
    For TempoTel = 0 To TxtInfo.UBound
        If InStr(TxtInfo(TempoTel).Tag, "|") <> 0 Then
            TxtInfo(TempoTel).text = Dec(Val(TxtInfo(TempoTel).text) / EURO, "#######0.0000")
        End If
    Next
End If
VernieuwTicketPrijs
VernieuwStockWaarde

End Sub

Private Sub cmdTonen_Click()

If Trim$(txtLink.text) = "" Then
    Exit Sub
End If

'Dim frmB As New frmMiniBrowser
Dim r As Long

If Err Then MsgBox Error: Exit Sub
Dim strAdress As String
strAdress = txtLink.text
r = ShellExecute(0, "open", strAdress, 0, 0, 1)
        
'frmB.WindowState = vbNormal
'If Err Then MsgBox Error: Exit Sub
'frmB.Show
'If Err Then MsgBox Error: Exit Sub
   
End Sub

Private Sub CmdVerwijderFiche_Click()

If Trim$(TLB_RECORD(TABLE_PRODUCTS)) = "" Or Trim$(TxtInfo(0)) = "" Then
    Msg = "Nummer met blanco verwijderen.  Bent u zeker"
    KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2)
    If KtrlBox = vbYes Then
        bDelete TABLE_PRODUCTS
        CmdHoger_Click
        Exit Sub
    End If
Else
    Msg = "Gegevens bestaande '" + bstNaam(TABLE_PRODUCTS) + "'-fiche :" + Trim$(TxtInfo(0)) + " verwijderen.  Bent U zeker ?"
    KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2)
    If KtrlBox = vbYes Then
        bDelete TABLE_PRODUCTS
        CmdHoger_Click
        Exit Sub
    End If
End If

End Sub

Private Sub cmdSchoon_Click()

    On Error Resume Next
    Schoon
    TxtInfo(0).SetFocus

End Sub

Private Sub Form_Load()

Set rsSQLQuery = New ADODB.Recordset
rsSQLQuery.CursorLocation = adUseClient
Set rsJourQuery = New ADODB.Recordset
rsJourQuery.CursorLocation = adUseClient

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Dim RekenOpties As String
cmdSwitch.Caption = "Ingave in EUR"

Schoon
RekenOpties = String99(READING, 181)
For COUNT_TO = 1 To Len(RekenOpties)
    If Mid(RekenOpties, COUNT_TO, 1) = 1 Then
        chkFilter(COUNT_TO - 1).Value = 1
    Else
        chkFilter(COUNT_TO - 1).Value = 0
    End If
Next
groepenVullen

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
rsSQLQuery.Close
Set rsSQLQuery = Nothing

End Sub

Private Sub Groepen_Click()

    ProductGroepen.Show 1
    groepenVullen
    
End Sub

Private Sub LijstRap_Click()

FrmLijstProdukten.Show

End Sub


Private Sub TxtInfo_GotFocus(Index As Integer)

IsGewijzigd = False
TxtInfo(Index).SelStart = 0
TxtInfo(Index).SelLength = Len(TxtInfo(Index).text)
iTabIndex = Index

TxtInfo(Index).BackColor = &HFFFF80
If InStr(TxtInfo(Index).Tag, ";") Then
    If RTrim$(TxtInfo(Index).text) = "" Then
        SnelHelpPrint "Druk [Ctrl] om te kiezen", BL_LOGGING
    ElseIf InStr(TxtInfo(Index).Tag, "&") Then
        Select Case Mid(TxtInfo(Index).Tag, InStr(TxtInfo(Index).Tag, "&") + 1, 1)
            Case "K"
                SharedFl = TABLE_CUSTOMERS
            Case "L"
                SharedFl = TABLE_SUPPLIERS
            Case "R"
                SharedFl = TABLE_LEDGERACCOUNTS
            Case Else
                MsgBox "nog niks"
        End Select
        bGet SharedFl, 0, TxtInfo(Index).text
        If Ktrl Then
            MsgBox TxtInfo(Index).text + " bestaat niet (meer) !"
        Else
            RecordToVeld SharedFl
            SnelHelpPrint FVT(SharedFl, 1) + " Druk [Ctrl] om te wijzigen", BL_LOGGING
        End If
    Else
        SnelHelpPrint "Druk [Ctrl] om te wijzigen", BL_LOGGING
    End If
End If

End Sub

Private Sub TxtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
Dim DummyText As String

If KeyCode <> 17 Then Exit Sub

If InStr(TxtInfo(Index).Tag, "&") Then
    aIndex = 0
    Select Case Mid(TxtInfo(Index).Tag, InStr(TxtInfo(Index).Tag, "&") + 1, 1)
        Case "K"
            SharedFl = TABLE_CUSTOMERS
        Case "L"
            SharedFl = TABLE_SUPPLIERS
        Case "R"
            SharedFl = TABLE_LEDGERACCOUNTS
        Case Else
            MsgBox "nog niks"
    End Select
    GridText = TxtInfo(Index).text
    SqlSearch.Show 1
    If Ktrl = 0 Then
        TxtInfo(Index).text = FVT(SharedFl, 0)
    End If
ElseIf InStr(TxtInfo(Index).Tag, ";") Then
    aIndex = Val(Mid(TxtInfo(Index).Tag, InStr(TxtInfo(Index).Tag, ";") + 1))
    aIndex = aIndex + 1000
    DummyText = TxtInfo(Index).text
    GridText = DummyText
    KeuzeVSF.Show 1
    If GridText <> DummyText Then
        DummyText = GridText
        TxtInfo(Index).text = DummyText
        VernieuwTicketPrijs
    End If
End If

End Sub

Private Sub TxtInfo_KeyPress(Index As Integer, KeyAscii As Integer)

If InStr(TxtInfo(Index).Tag, ";") Then
    KeyAscii = 0
End If
IsGewijzigd = True

End Sub


Private Sub TxtInfo_LostFocus(Index As Integer)
Dim BedragZonder As Double
Dim Sleuteltje As String

TxtInfo(Index).BackColor = &HFFFFFF
Select Case Index
    Case 0
        bGet TABLE_PRODUCTS, 0, TxtInfo(0).text
        If Ktrl Then
            Sleuteltje = TxtInfo(0).text
            Schoon
            FicheIsNieuw = True
            TxtInfo(0).text = Sleuteltje
        Else
            VensterVullen
        End If
        
    Case 1
        If Not IsGewijzigd Then
            Exit Sub
        Else
            frmProduktFiche.Caption = "ProduktFiche : " + TxtInfo(1)
        End If
    
    Case 3, 6 'Rekenwaarden zijn gewijzigd
        
        Select Case Index
            Case 3
                If Val(TxtInfo(3)) = 0 Then TxtInfo(3) = "1"
                TxtInfo(3) = Dec$(Val(TxtInfo(3)), "#####.00")
            Case 6
                TxtInfo(6) = Dec$(Val(TxtInfo(6)), "###")
                If chkFilter(1) = 1 Then
                    If chkFilter(0) = 1 Then
                        BedragZonder = Val(TxtInfo(9))
                    Else
                        BedragZonder = Val(TxtInfo(9)) * Val(TxtInfo(3))
                    End If
                    TxtInfo(8) = Dec$(BedragZonder + (BedragZonder * Val(TxtInfo(6)) / 100), "########0.00000")
                    'TxtInfo(8) = (BedragZonder + (BedragZonder * Val(TxtInfo(6)) / 100))
                End If
        End Select
        VernieuwTicketPrijs
        
    Case 8 'bedragen eerst kontroleren op
           'verkoop  IN/EX + GLOBAAL
        If Not IsGewijzigd Then Exit Sub
        If chkFilter(3) = 1 Then
            TxtInfo(8) = Dec$(Val(TxtInfo(8)) * 100 / (100 + Val(Mid(TxtInfo(7), InStr(TxtInfo(7), ":") + 1))), "########0.00000")
        End If
        If chkFilter(0) = 1 Then
            TxtInfo(8) = Dec$(Val(TxtInfo(8)) / Val(TxtInfo(3)), "########0.00000")
        End If
        TxtInfo(8) = Dec$(Val(TxtInfo(8)), "########0.00000")
        VernieuwTicketPrijs
    
    Case 9 'bedragen eerst kontroleren op
           'aankoop IN/EX + GLOBAAL
        If Not IsGewijzigd Then Exit Sub
        If chkFilter(2) = 1 Then
            TxtInfo(9) = Dec$(Val(TxtInfo(9)) * 100 / (100 + Val(Mid(TxtInfo(7), InStr(TxtInfo(7), ":") + 1))), "########0.00000")
        End If
        If chkFilter(0) = 1 Then
            TxtInfo(9) = Dec$(Val(TxtInfo(9)) / Val(TxtInfo(3)), "########0.00000")
        End If
        If chkFilter(1) = 1 Then
            If chkFilter(0) = 1 Then
                BedragZonder = Val(TxtInfo(9))
            Else
                BedragZonder = Val(TxtInfo(9)) * Val(TxtInfo(3))
            End If
            TxtInfo(8) = Dec$(BedragZonder + (BedragZonder * Val(TxtInfo(6)) / 100), "########0.00000")
        End If
        TxtInfo(9) = Dec$(Val(TxtInfo(9)), "########0.00000")
        VernieuwTicketPrijs
        VernieuwStockWaarde
           
    Case 14 To 16
        TxtInfo(Index) = Dec$(Val(TxtInfo(Index)), "#####0.000")
        VernieuwStockWaarde
        
    Case 17 To 19   'stockcijfers zijn gewijzigd
        VernieuwStockWaarde
        
End Select

End Sub




Private Sub v_Click(PreviousTab As Integer)

If v.Caption = "Journaal" Then
    Screen.MousePointer = vbHourglass
    
    Dim msgQuery As String

    msgQuery = "SELECT * FROM Journalen WHERE v102 ='" & vBibTekst(TABLE_PRODUCTS, "v102") & "' and v019 >'5'"
    On Error Resume Next
    msfJournaal.Recordset = Nothing
    rsJourQuery.Close
    On Error Resume Next
    rsJourQuery.Open msgQuery, adntDB
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        msfSQL.Refresh
    Else
         Set msfJournaal.DataSource = rsJourQuery
    End If
    lbJournaal = rsJourQuery.RecordCount
    Screen.MousePointer = vbNormal
End If

End Sub

Private Sub VerwijderenMogelijk_Click()
        
VerwijderenMogelijk.Checked = Not VerwijderenMogelijk.Checked
If VerwijderenMogelijk.Checked Then
    CmdVerwijderFiche.Enabled = True
Else
    CmdVerwijderFiche.Enabled = False
End If

End Sub


Private Sub groepenVullen()

    Dim splitstring() As String
    Dim COUNT_TOHier As Integer
    
    Set rsGroepenHier = New ADODB.Recordset
    rsGroepenHier.CursorLocation = adUseServer
    
    On Error Resume Next
    Err = 0
    rsGroepenHier.Open "p_Groepen", adntDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect '  adLockReadOnly, adCmdTableDirect
    If Err Then
        MsgBox "productgroepen worden hierna geïnitialiseerd"
        Groepen_Click
        Exit Sub
    End If
    rsGroepenHier.Index = "GroepsNaam"
    
    On Error Resume Next
    Me.cbCategorie.Clear
    Err = 0
    rsGroepenHier.Seek "Categorie", adSeekFirstEQ
    If Err Then
        MsgBox "Gelieve de groepen te initialiseren a.u.b."
    Else
        'lijstje samenstellen
        splitstring = Split(rsGroepenHier.fields("GroepItems"), ";")
        For COUNT_TOHier = 0 To UBound(splitstring)
            Me.cbCategorie.AddItem splitstring(COUNT_TOHier)
        Next
    End If
    Me.cbMerk.Clear
    Err = 0
    rsGroepenHier.Seek "Merk", adSeekFirstEQ
    If Err Then
        MsgBox "Gelieve de groepen te initialiseren a.u.b."
    Else
        'lijstje samenstellen
        splitstring = Split(rsGroepenHier.fields("GroepItems"), ";")
        For COUNT_TOHier = 0 To UBound(splitstring)
            Me.cbMerk.AddItem splitstring(COUNT_TOHier)
        Next
    End If
    rsGroepenHier.Close
    Set rsGroepenHier = Nothing
    
End Sub
