VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form DirekteAankoop 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7770
   ClientLeft      =   1170
   ClientTop       =   1365
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "0200001.frx":0000
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7770
   ScaleWidth      =   8070
   Visible         =   0   'False
   Begin VB.TextBox TextBoxWarningTestCompany 
      BackColor       =   &H000000FF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2160
      TabIndex        =   43
      TabStop         =   0   'False
      Text            =   "Opgelet: Dit is een testbedrijf!"
      Top             =   3120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   7695
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   13573
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Ingave"
      TabPicture(0)   =   "0200001.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "OMSLabel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(7)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(14)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(8)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(12)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(11)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LabelDocumentReference"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TekstInfo12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TekstInfo5"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TekstInfo7"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TekstInfo9"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TekstInfo0"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TekstInfo1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TekstInfo2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TekstInfo10"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TekstInfo6"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TekstInfo3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TekstInfo4"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "AankoopDetail"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "AankoopOptie(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "AankoopOptie(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "AankoopOptie(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdSQLInfo"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "StockBeheer"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Annuleren"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Medekontraktant"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Afsluiten"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Kontrole"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Schoonvegen"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "SSTab1"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "cbImportUBL"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cbCheckTools"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "ButtonOptimize"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "CheckBoxAlwaysPeppolRefresh"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).ControlCount=   38
      TabCaption(1)   =   "Bijlage"
      TabPicture(1)   =   "0200001.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbAfbeelding"
      Tab(1).Control(1)=   "OLE1"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Voorkeurinstellingen"
      TabPicture(2)   =   "0200001.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "obTAB(1)"
      Tab(2).Control(1)=   "obTAB(0)"
      Tab(2).Control(2)=   "Label3"
      Tab(2).ControlCount=   3
      Begin VB.CheckBox CheckBoxAlwaysPeppolRefresh 
         Alignment       =   1  'Right Justify
         Caption         =   "Check Peppol"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   0
         Width           =   1455
      End
      Begin VB.CommandButton ButtonOptimize 
         Caption         =   "Dubbels Vermijden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6240
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   5160
         Width           =   1635
      End
      Begin VB.CommandButton cbCheckTools 
         Caption         =   "Check Tools"
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
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1425
      End
      Begin VB.CommandButton cbImportUBL 
         Caption         =   "&UBL B2B IN"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5040
         TabIndex        =   0
         TabStop         =   0   'False
         ToolTipText     =   "B2B Aankoopdocumenten Monitor"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.OptionButton obTAB 
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
         Height          =   195
         Index           =   1
         Left            =   -74760
         TabIndex        =   38
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton obTAB 
         Caption         =   "Ingave factuur"
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
         Index           =   0
         Left            =   -74760
         TabIndex        =   37
         Top             =   480
         Width           =   1575
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   1875
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   3307
         _Version        =   393216
         Style           =   1
         Tab             =   1
         TabHeight       =   494
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "&Leverancier"
         TabPicture(0)   =   "0200001.frx":035E
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "LeverancierInfo"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Bewerken"
         TabPicture(1)   =   "0200001.frx":037A
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "LabelInfoXlog"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "cmdXLog"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "&Historiek"
         TabPicture(2)   =   "0200001.frx":0396
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "ListView1"
         Tab(2).ControlCount=   1
         Begin VB.CommandButton cmdXLog 
            Caption         =   "Geactiveerde leverancier Inkijken / Wijzigen"
            Enabled         =   0   'False
            Height          =   495
            Left            =   120
            TabIndex        =   40
            Top             =   1200
            Width           =   4575
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1395
            Left            =   -74880
            TabIndex        =   41
            Top             =   360
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   2461
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            HotTracking     =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483625
            BackColor       =   -2147483624
            BorderStyle     =   1
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
            NumItems        =   0
         End
         Begin VB.Label LabelInfoXlog 
            Alignment       =   2  'Center
            Caption         =   "Voor manuele verrichting activeer een leverancier via TAB 'Leverancier' of TAB 'Historiek'"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   42
            Top             =   480
            Width           =   4575
         End
         Begin VB.Label LeverancierInfo 
            BackColor       =   &H80000018&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1395
            Left            =   -74880
            TabIndex        =   30
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.CommandButton cbAfbeelding 
         Caption         =   "Openen"
         Height          =   375
         Left            =   -74280
         TabIndex        =   34
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Schoonvegen 
         Caption         =   "Sch&oon"
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
         Left            =   5040
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   780
         Width           =   1215
      End
      Begin VB.CommandButton Kontrole 
         Caption         =   "&Controle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6240
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1635
      End
      Begin VB.CommandButton Afsluiten 
         Caption         =   "&Boeken"
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
         Height          =   360
         Left            =   6240
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   5880
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CheckBox Medekontraktant 
         Alignment       =   1  'Right Justify
         Caption         =   "&Medecontractant"
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
         Height          =   240
         Left            =   2640
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1575
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
         Height          =   315
         Left            =   5040
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1140
         Width           =   1215
      End
      Begin VB.CheckBox StockBeheer 
         Caption         =   "&Stockbeheer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4320
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdSQLInfo 
         Caption         =   "SQL &Info"
         Height          =   315
         Left            =   3840
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.OptionButton AankoopOptie 
         Alignment       =   1  'Right Justify
         Caption         =   "&Facturering"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   6660
         TabIndex        =   2
         Top             =   720
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.OptionButton AankoopOptie 
         Alignment       =   1  'Right Justify
         Caption         =   "CreditNota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   6720
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   960
         Width           =   1065
      End
      Begin VB.OptionButton AankoopOptie 
         Alignment       =   1  'Right Justify
         Caption         =   "Leveringsbon"
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
         Height          =   240
         Index           =   2
         Left            =   6480
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1200
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.ListBox AankoopDetail 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   2760
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   7845
      End
      Begin MSMask.MaskEdBox TekstInfo4 
         Height          =   315
         Left            =   1680
         TabIndex        =   11
         Top             =   6840
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   -2147483624
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         MaxLength       =   20
         Mask            =   "+++###/####/#####+++"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo3 
         Height          =   315
         Left            =   4980
         TabIndex        =   28
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColor       =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo6 
         Height          =   315
         Left            =   4560
         TabIndex        =   17
         Top             =   5880
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColor       =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo10 
         Height          =   315
         Left            =   6720
         TabIndex        =   29
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColor       =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo2 
         Height          =   315
         Left            =   1680
         TabIndex        =   9
         Top             =   5880
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColor       =   0
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo1 
         Height          =   315
         Left            =   1680
         TabIndex        =   7
         Top             =   5160
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColor       =   0
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo0 
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   5520
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColor       =   0
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo9 
         Height          =   315
         Left            =   6360
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   6840
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColor       =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo7 
         Height          =   315
         Left            =   4560
         TabIndex        =   20
         Top             =   5520
         Visible         =   0   'False
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColor       =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo5 
         Height          =   315
         Left            =   4560
         TabIndex        =   14
         Top             =   5160
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColor       =   0
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo12 
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Top             =   6360
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   556
         _Version        =   393216
         ClipMode        =   1
         BackColor       =   -2147483624
         ForeColor       =   0
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         PromptChar      =   "_"
      End
      Begin VB.Label LabelDocumentReference 
         Alignment       =   1  'Right Justify
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
         Left            =   3600
         TabIndex        =   48
         Top             =   7440
         Width           =   4335
      End
      Begin VB.Label Label1 
         Caption         =   "DocID Leveranc."
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   45
         Top             =   6435
         Width           =   1515
      End
      Begin VB.Label Label3 
         Caption         =   "Tabblad bij opstart"
         Height          =   255
         Left            =   -74760
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
      Begin VB.OLE OLE1 
         AutoActivate    =   0  'Manual
         Class           =   "AcroExch.Document.7"
         Height          =   4095
         Left            =   -72840
         SizeMode        =   1  'Stretch
         TabIndex        =   35
         Top             =   120
         Width           =   3015
      End
      Begin VB.Label Label1 
         Caption         =   "Totaal"
         Height          =   195
         Index           =   11
         Left            =   3000
         TabIndex        =   16
         Top             =   6000
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Btw Verschuldigd"
         Height          =   240
         Index           =   12
         Left            =   3000
         TabIndex        =   15
         Top             =   5595
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Btw Aftrekbaar"
         Height          =   240
         Index           =   8
         Left            =   3000
         TabIndex        =   13
         Top             =   5235
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Betaalreferte"
         Height          =   240
         Index           =   14
         Left            =   120
         TabIndex        =   12
         Top             =   6840
         Width           =   1395
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Default&Rek. (Leverancier en Btw)"
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
         Left            =   4980
         TabIndex        =   27
         Top             =   60
         Width           =   2835
      End
      Begin VB.Label Label1 
         Caption         =   "Vervaldag"
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   33
         Top             =   5955
         Width           =   1035
      End
      Begin VB.Label OMSLabel 
         Caption         =   "&Alg.Rek Naam rekening                                Totaal    Artikelnr.    Omschrijving"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   1920
         Width           =   9465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Koers"
         Height          =   240
         Index           =   3
         Left            =   6480
         TabIndex        =   31
         Top             =   6480
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Documentdatum      =  Boekdatum"
         Height          =   675
         Left            =   120
         TabIndex        =   6
         Top             =   5160
         Width           =   1485
      End
   End
End
Attribute VB_Name = "DirekteAankoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim documentRefIsOGM As Boolean
Dim Document_Key As String * 11
Dim Document_OGM As String
Dim Document_OGM_NoFormat As String

Dim adminNoVat As Boolean

Dim ifSupplierInsertWarning As Boolean

Dim rsDocuments As New ADODB.Recordset
Dim rsSupplier As New ADODB.Recordset

Dim supplierA110 As String

Dim AankoopDetailTitel(8) As String
Dim lvLeverancier As String
Dim GeScanBestand As String


Dim PriveRekening As String * 7
Dim LeverancierRekening As String * 7
Dim GrensDetail(3) As String * 14

Dim AankoopFlg As Integer
Dim rbtwVAK(10) As String * 7
Dim sMuntLever As String * 3
Dim sIsIntraFlg As String * 1

Dim StartBlad As Boolean

Dim Ar As Integer

Dim VeldRekening As String * 7
Dim VeldNaam            As String * 40
Dim VeldBedrag          As String * 10

Dim supplierCompanyId As String
Dim supplierCountryCode As String
Dim supplierVatNumber As String

Dim Positie As Integer

Function CheckDocument() As String

    CheckDocument = ""
    'For T = 0 To UBound(uitwisselingOMSArray): Print T; " "; uitwisselingOMSArray(T); ": "; uitwisselingDATAArray(T): Next
            
    Set rsDocuments = New ADODB.Recordset
    On Error Resume Next
    Err = 0
    rsDocuments.CursorLocation = adUseClient
    
    Dim payRefClean As String
    
    payRefClean = Replace(uitwisselingDATAArray(14), "/", "")
    payRefClean = Replace(payRefClean, "+", "")
    
    'If payRefClean = "" Then
    '    payRefClean = uitwisselingDATAArray(1)
    'End If
            
    Dim sSQL As String
    sSQL = "SELECT "
    sSQL = sSQL + "Leveranciers.A110, Leveranciers.v404, "
    sSQL = sSQL + "Dokumenten.v409, Dokumenten.v033, Dokumenten.v034, Dokumenten.v035, Dokumenten.v036, Dokumenten.v037, Dokumenten.v039, Dokumenten.v249 "
    sSQL = sSQL + "FROM  Leveranciers, Dokumenten "
    sSQL = sSQL + "WHERE Leveranciers.v404 = '" + uitwisselingDATAArray(5) + "' "
    sSQL = sSQL + "AND Dokumenten.v034 = 'L'+Leveranciers.A110 "
    sSQL = sSQL + "AND Dokumenten.v409 = '" + uitwisselingDATAArray(1) + "' "
    sSQL = sSQL + "AND Dokumenten.v039 = '" + payRefClean + "'"
    
    SnelHelpPrint Msg, BL_LOGGING
    Screen.MousePointer = vbHourglass
    rsDocuments.Open sSQL, adntDB, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        Exit Function
    ElseIf rsDocuments.RecordCount = 0 Then
        CheckDocument = "In te boeken"
        Exit Function
    ElseIf rsDocuments.RecordCount = 1 Then
        'check if already booked
        XLogKey = ""
        If (Val(rsDocuments("v249")) = Val(uitwisselingDATAArray(21))) And (payRefClean = Trim(rsDocuments("v039"))) And (Replace(uitwisselingDATAArray(2), "-", "") = rsDocuments("v035")) Then
            'Msg = "Dit aankoopdocument is reeds ingeboekt." & vbCrLf & vbCrLf
            'Msg = Msg & "Document ID Leverancier: " & documentIdToCheck & vbCrLf & vbCrLf
            'Msg = Msg & "Betaalreferte : " & payReferenceToCheck & vbCrLf
            'Msg = Msg & "Totaal Bedrag : " & (Val(totalInclusiveVATToCheck)) & vbCrLf
            'Msg = Msg & "Datum Document: " & DATE_TEXT(rsDocuments("v035")) & vbCrLf & vbCrLf
            'Msg = Msg & "MarSync zal automatisch na betaling procedure archiveren opstarten"
            'MsgBox Msg, vbInformation
            'Schoon
            
            CheckDocument = "Ingeboekt"
            XLogKey = rsDocuments("v033")
            If IsNull(rsDocuments("v037")) Then
                CheckDocument = CheckDocument & "Te betalen"
                Exit Function
            ElseIf (Val(rsDocuments("v249")) - Val(rsDocuments("v037"))) = 0 Then
                CheckDocument = CheckDocument & "Betaald"
                Exit Function
            Else
                CheckDocument = CheckDocument & "Gedeeltelijk openstaand"
                Exit Function
            End If
            CheckDocument = "another strange situation"
            Exit Function
        Else
            CheckDocument = "whatsgoingon?"
            Exit Function
        End If
    ElseIf rsDocuments.RecordCount > 1 Then
        MsgBox "Meerdere documenten gevonden, onlogische situatie.", vbExclamation
        Exit Function
    End If
    
End Function

Function CheckSupplier() As Boolean
       
    CheckSupplier = False
    Set rsSupplier = New ADODB.Recordset
    
    On Error Resume Next
    Err = 0
    rsSupplier.CursorLocation = adUseClient
    
    Dim searchOnVat As String
    Dim searchCountry As String
    
    uitwisselingOMSArray() = Split(uitwisselingOMS, vbTab)
    uitwisselingDATAArray() = Split(uitwisselingDATA, vbTab)
    'For T = 0 To UBound(uitwisselingOMSArray): Print T; " "; uitwisselingOMSArray(T); ": "; uitwisselingDATAArray(T): Next
    '00 UBL VersionID: 2.1
    '01 documentIdToCheck:   V0202550406 -220935
    '02 dateSellerDocumentToCheck: 2025-11-08
    '03 dateExpiringDocumentToCheck: 2025-11-08
    '04 invoiceTypeCode: 380
    '05 supplierCompanyIdToCheck: 0440058217
    '06 supplierID:
    '07 supplierNameToCheck: Bv Zakenkantoor Hedwig Roelandt En Jos Vermoesen
    '08 supplierStreetToCheck: Grote Baan 141
    '09 supplierCityToCheck:           Herdersem
    '10 supplierPostalCodeToCheck: 9310
    '11 supplierCountryCodeToCheck:   BE
    '12 supplierVatNumberToCheck:   BE0440058217
    '13 paymentMeansCode: 1
    '14 payReferenceToCheck:   V0202550406 -220935
    '15 supplierIBANToCheck:   BE83891854037015
    '16 paySupplierNameToCheck:
    '17 supplierBICToCheck:   VDSPBE91
    '18 supplierDomMandate:
    '19 supplierDomDate:
    '20 totalExclusiveVAT: 110.37
    '21 totalInclusiveVAT: 133.55
    '22 customerAccountID:
    '23 customerEndpointID: 0874419267
    '24 customerName: Vermoesen Beheer CommV
    '25 customerStreetName: Grote Baan 141
    '26 customerCityName:   Herdersem
    '27 customerPostalZone: 9310
    '28 customerCountryCode:   BE
    '29 customerTaxID:           BE0874419267
    '30 customerTaxScheme:   VAT
    
    searchCountry = uitwisselingDATAArray(11)
    If searchCountry = "BE" Then
        searchOnVat = uitwisselingDATAArray(5)
    ElseIf searchCountry = "NL" Then
        searchOnVat = Mid(uitwisselingDATAArray(12), 3)
    Else
        MsgBox "Nog logica te voorzien voor " & uitwisselingDATAArray(7), vbInformation
        searchOnVat = Mid(uitwisselingDATAArray(12), 3)
    End If
    'check with supplierCompanyId OR with supplierVatNumber
    'v404 or A161
    
    Msg = "SELECT A110, v150, A161, v404, A100, v259 FROM Leveranciers WHERE v150 = '" & searchCountry & "' AND A161 = '" & searchOnVat & "'"
    SnelHelpPrint Msg, BL_LOGGING
    Screen.MousePointer = vbHourglass
    rsSupplier.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        Exit Function
    ElseIf rsSupplier.RecordCount = 1 Then
        '15 supplierIBANToCheck:   BE83891854037015
        'vBib TABLE_SUPPLIERS, uitwisselingDATAArray(15), "v259"  'IBAN
        If Trim(uitwisselingDATAArray(15)) = "" Then
        ElseIf Trim(uitwisselingDATAArray(15)) = objectValue(rsSupplier("v259")) Then
        Else
            Msg = "IBAN rekening is verschillend met database." & vbCrLf & vbCrLf
            Msg = Msg & "* Peppol document: " & uitwisselingDATAArray(15) & vbCrLf
            Msg = Msg & "* Tabel leveranciers: " & rsSupplier("v259") & vbCrLf & vbCrLf
            Msg = Msg & "Wijzig de leveranciersfiche manueel indien nodig."
            MsgBox Msg, vbExclamation, "IBAN verschillend"
        End If
        
    ElseIf rsSupplier.RecordCount = 0 Then
        If ifSupplierInsertWarning Then
            Msg = "Geen leverancier gevonden met het ondernemingsnummer " & uitwisselingDATAArray(5) & vbCrLf & vbCrLf
            Msg = Msg & "Naam: " & uitwisselingDATAArray(7) & vbCrLf
            Msg = Msg & "Straat: " & uitwisselingDATAArray(8) & vbCrLf
            Msg = Msg & "Plaatsnaam: " & uitwisselingDATAArray(9) & vbCrLf & vbCrLf
            Msg = Msg & "Controleer eerst of voeg een nieuwe leverancier toe" & vbCrLf & vbCrLf
            Msg = Msg & "Hierna wordt voorgesteld een nieuwe leverancier aan te maken" & vbCrLf & "met gegevens uit het UBL document en ondernemingsnummer als codenummer." & vbCrLf & "Wijzig het codenummer indien nodig en druk enter om de fiche te controleren, te vervolledigen." & vbCrLf & vbCrLf
            
            MsgBox Msg, vbInformation, "KBO nummer: " & uitwisselingDATAArray(5)
               
            On Error Resume Next
            INSERT_FLAG(TABLE_SUPPLIERS) = 1
                
            PeppolFlag = True
            TLB_RECORD(TABLE_SUPPLIERS) = ""
            vBib TABLE_SUPPLIERS, "2", "A10C"    'Taalkode
            vBib TABLE_SUPPLIERS, searchCountry, "v150"   'Landkode    ISO kode
            vBib TABLE_SUPPLIERS, "002", "v149"  'Landnummer  ISO kode
            vBib TABLE_SUPPLIERS, "B  ", "A109"  'Landkode Postkantoor
            vBib TABLE_SUPPLIERS, "EUR", "vs03"  'Munteenheid ISO kode
            vBib TABLE_SUPPLIERS, "1", "vs07"    'exemplaren dokumenten
        
            vBib TABLE_SUPPLIERS, uitwisselingDATAArray(5), "A110"  'Uniek Codenummer
            vBib TABLE_SUPPLIERS, uitwisselingDATAArray(5), "v404"  'Ondernemingsnummer
                        
            If searchCountry = "BE" Then
                vBib TABLE_SUPPLIERS, uitwisselingDATAArray(5), "A161"  'Ondernemingsnummer = Btw nummer
            Else
                If uitwisselingDATAArray(12) = "" Then
                    vBib TABLE_SUPPLIERS, uitwisselingDATAArray(5), "A161"  'Ondernemingsnummer
                Else
                    vBib TABLE_SUPPLIERS, Mid(uitwisselingDATAArray(12), 3), "A161" 'BTW nummer
                End If
            End If
                
            vBib TABLE_SUPPLIERS, uitwisselingDATAArray(7), "A100"  'Naam
            vBib TABLE_SUPPLIERS, uitwisselingDATAArray(8), "A104"  'Straat en nummer
            vBib TABLE_SUPPLIERS, uitwisselingDATAArray(9), "A108"  'Plaats
            vBib TABLE_SUPPLIERS, uitwisselingDATAArray(10), "A107"  'Postcode
            vBib TABLE_SUPPLIERS, uitwisselingDATAArray(15), "v259"  'IBAN
            vBib TABLE_SUPPLIERS, uitwisselingDATAArray(17), "v260"  'BIC
            'vBib TABLE_SUPPLIERS, supplierCompanyIdToCheck, "A110"  'Uniek Codenummer

            BasisB(TABLE_SUPPLIERS).WindowState = 0
            BasisB(TABLE_SUPPLIERS).SetFocus
            BasisB(TABLE_SUPPLIERS).TekstInfo(0).text = uitwisselingDATAArray(5)
            BasisB(TABLE_SUPPLIERS).Knop(5).Enabled = True
            BasisB(TABLE_SUPPLIERS).Knop(5).Default = True
            BasisB(TABLE_SUPPLIERS).TekstInfo(0).Enabled = True
            BasisB(TABLE_SUPPLIERS).TekstInfo(0).SetFocus
            Exit Function
        End If
    ElseIf rsSupplier.RecordCount > 1 Then
        Msg = "Meerdere leveranciers gevonden met ondernemingsnummer " & searchOnVat & ". Verbeter eerst" & vbCrLf & vbCrLf
        rsSupplier.MoveFirst
        Msg = Msg & "=> " & Trim(rsSupplier("A110")) & " " & Trim(rsSupplier("A100")) & vbCrLf
        Do While Not rsSupplier.EOF
            rsSupplier.MoveNext
            Msg = Msg & "=> " & Trim(rsSupplier("A110")) & " " & Trim(rsSupplier("A100")) & vbCrLf
        Loop
        MsgBox Msg, vbExclamation, "KBO nummer: " & uitwisselingDATAArray(5)
        Exit Function
    End If
    'If searchCountry = "BE" Then
    'Else
    '    Stop
    'End If
    CheckSupplier = True

End Function

Function CheckSupplierDocuments(supplierCode As String) As Boolean
       
    Dim valueInRecordV410 As String
    Dim checkWithVatNumber As String
    
    CheckSupplierDocuments = False
    
    Dim rsAny As ADODB.Recordset
    Set rsAny = New ADODB.Recordset
    
    On Error Resume Next
    Err = 0
    rsAny.CursorLocation = adUseClient
    
    Msg = "SELECT A110, v150, A161, v404, v410, A100 FROM Leveranciers WHERE A110 = '" & supplierCode & "'"
    SnelHelpPrint Msg, BL_LOGGING
    Screen.MousePointer = vbHourglass
    rsAny.Open Msg, adntDB, adOpenDynamic, adLockOptimistic
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    ElseIf rsAny.RecordCount = 1 Then
        If IsNull(rsAny("v410")) Or rsAny("v410").Value = "" Then
            'First check with 0208:"
            checkWithVatNumber = "0208:" + rsAny("V404")
            valueInRecordV410 = CheckPeppolRegistration(checkWithVatNumber)
            'MsgBox "lengte: " + Str(Len(valueInRecordV410))
            'TODO: definitly to check for Peppol readiness
            If Len(valueInRecordV410) < 300 Then
                Msg = "Gecontroleerd met code: "
                Msg = Msg & checkWithVatNumber & vbCrLf & vbCrLf
                Msg = Msg & "Mogelijk geen Peppol Registratie" & vbCrLf & vbCrLf
                Msg = Msg & "Tot slot controleren met verouderde 9925:BE"
                MsgBox Msg, vbInformation
                checkWithVatNumber = "9925:BE" + rsAny("V404")
                valueInRecordV410 = CheckPeppolRegistration(checkWithVatNumber)
                'MsgBox "lengte: " + Str(Len(valueInRecordV410))
                'TODO: definitly to check for Peppol readiness
                If Len(valueInRecordV410) < 300 Then
                    CheckSupplierDocuments = False
                Else
                    CheckSupplierDocuments = True
                End If
            Else
                CheckSupplierDocuments = True
            End If
            rsAny("v410") = valueInRecordV410
            rsAny.Update
        ElseIf Len(rsAny("v410")) > 500 And Me.CheckBoxAlwaysPeppolRefresh.Value = vbUnchecked Then
            CheckSupplierDocuments = True
        ElseIf Me.CheckBoxAlwaysPeppolRefresh.Value = vbChecked Then 'always refresh!
            'First check with 0208:"
            checkWithVatNumber = "0208:" + rsAny("V404")
            valueInRecordV410 = CheckPeppolRegistration(checkWithVatNumber)
            'MsgBox "lengte: " + Str(Len(valueInRecordV410))
            'TODO: definitly to check for Peppol readiness
            If Len(valueInRecordV410) < 300 Then
                Msg = "Gecontroleerd met code: "
                Msg = Msg & checkWithVatNumber & vbCrLf & vbCrLf
                Msg = Msg & "Mogelijk geen Peppol Registratie" & vbCrLf & vbCrLf
                Msg = Msg & "Tot slot controleren met verouderde 9925:BE"
                MsgBox Msg, vbInformation
                checkWithVatNumber = "9925:BE" + rsAny("V404")
                valueInRecordV410 = CheckPeppolRegistration(checkWithVatNumber)
                'MsgBox "lengte: " + Str(Len(valueInRecordV410))
                'TODO: definitly to check for Peppol readiness
                If Len(valueInRecordV410) < 300 Then
                    CheckSupplierDocuments = False
                Else
                    CheckSupplierDocuments = True
                End If
            Else
                CheckSupplierDocuments = True
            End If
            rsAny("v410") = valueInRecordV410
            rsAny.Update
        End If
        rsAny.Close
    End If
    Set rsAny = Nothing
    
End Function


Private Sub Aankoopdetail_DblClick()

If AankoopDetail.ListIndex = -1 Then
    AankoopDetail_KeyPress 43
Else
    AankoopDetail_KeyPress 13
End If

End Sub


Private Sub AankoopDetail_GotFocus()

SnelHelpPrint "[Insert] lijn bijvoegen, [Delete] lijn verwijderen, [spatie],[Enter] om te wijzigen", BL_LOGGING

On Error Resume Next
AankoopDetail.ListIndex = Positie

Kontrole.Default = False

End Sub


Private Sub AankoopDetail_KeyDown(KeyCode As Integer, Shift As Integer)

Dim X As Integer

Positie = AankoopDetail.ListIndex
Select Case KeyCode
    Case 45, 107
        GridText = ""
        Wijzigen.Show 1
        If GridText = "" Then
        Else
            If RekeningOK(Left(GridText, 7)) Then
                AankoopDetail.AddItem GridText, Positie + 1
                X = InvestKtrl()
            End If
        End If
    Case 46
        If Positie < 0 Then
            MsgBox "Eerst een lijn selecteren !", 0, "Lijn wijzigen"
            Exit Sub
        End If
        Msg = "Lijn verwijderen !  Bent U zeker ?"
        Ktrl = MsgBox(Msg$, 292)
        If Ktrl = 6 Then
            AankoopDetail.RemoveItem Positie
        End If
End Select
DirekteAankoop.AankoopDetail.SetFocus

End Sub

Private Sub AankoopDetail_KeyPress(KeyAscii As Integer)

Dim X       As Integer

On Local Error Resume Next
Positie = AankoopDetail.ListIndex
Select Case KeyAscii
    Case 13
        If Positie < 0 Then
            MsgBox "Eerst een lijn selecteren !", 0, "Lijn wijzigen"
            Exit Sub
        End If
        GridText = AankoopDetail.text
        Wijzigen.Show 1
        If GridText = "" Then
        Else
            If RekeningOK(Left(GridText, 7)) Then
                AankoopDetail.RemoveItem Positie
                AankoopDetail.AddItem GridText, Positie
                X = InvestKtrl()
                AankoopDetail.ListIndex = Positie
            End If
        End If
End Select
DirekteAankoop.AankoopDetail.SetFocus

End Sub






Private Sub AankoopOptie_Click(Index As Integer)

Jumper:
Select Case Index
    Case 0
        Ar = 1
        Document_Key = SleutelDok(Ar)
        RefreshReference

    Case 1
        Ar = 3
        Document_Key = SleutelDok(Ar)
        RefreshReference

    Case 2
        AankoopOptie(0).Value = True
End Select
DirekteAankoop.Caption = "Direkte aankoopverrichting         (" + Document_Key + ")"
bGet TABLE_INVOICES, 0, Document_Key
If Ktrl = 0 Then
    MsgBox "Document " & Document_Key & " is reeds aanwezig..." & vbCr & vbCr & "Controleer eventueel uw tellerbestand voor het active boekjaar.  Indien U zopas wisselde van boekjaar met het aankoopvenster open, mag U (na controle) deze melding negeren..." & vbCr & vbCr & "NIET BOEKJAAR GEWISSELD ZOPAS ?  EERST UW TELLEBESTAND + PROEF- SALDI BALANS CONTROLEREN !!!", vbExclamation
End If

End Sub

Private Sub Afsluiten_Click()
Dim A                   As String
Dim aa                  As String * 30
Dim dTTwb               As Double
Dim DummySleutel        As String

If AankoopDetail.ListCount = 0 Then Exit Sub

If Not DATE_CHECK((TekstInfo0.text), PERIODAS_TEXT) Then
    Beep
    On Error Resume Next
    TekstInfo0.SetFocus
    Exit Sub
End If

If Document_Key <> SleutelDok(Ar) Then
    Msg = Document_Key + " <> " + SleutelDok(Ar) + vbCrLf + vbCrLf
    Msg = Msg + "MOGELIJKHEID 1: Teller is identiek, boekjaar is hoger/lager." + vbCrLf
    Msg = Msg + "U hebt dus het actief boekjaar of periode gewijzigd tijdens de aanmaak van dit dokument.  Probeer nogmaals NA KONTROLE." + vbCrLf + vbCrLf
    Msg = Msg + "MOGELIJKHEID 2: Boekjaar is identiek, teller is hoger/lager." + vbCrLf
    Msg = Msg + "Een andere gebruiker heeft ondertussen een dokument verwerkt." + vbCrLf + vbCrLf
    Msg = Msg + "Kontroleer eerst eens vooraleer de boeking nogmaals uit te voeren a.u.b. !!!"
    MsgBox Msg
    Document_Key = SleutelDok(Ar)
    DirekteAankoop.Caption = "Direkte aankoopverrichting         (" + Document_Key + ")"
    Exit Sub
End If

Msg = "Document wegschrijven, boekhouding bijwerken."
Ktrl = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, Document_Key + " verwerken.")
Select Case Ktrl
    Case 7
        Exit Sub
End Select

Ktrl = bOpen(TABLE_LEDGERACCOUNTS)
Ktrl = bOpen(TABLE_JOURNAL)
Ktrl = bOpen(TABLE_INVOICES)
bBegin
If WegBoekFout() Then
    bAbort
    DirekteAankoop.SetFocus
    Exit Sub
Else
    bEnd
    If Ktrl Then
        bAbort
        Exit Sub
    End If
    If AankoopFlg = 1 And sIsIntraFlg = "1" Then
        Fl = TABLE_SUPPLIERS
        aIndex = 19
        dTTwb = Val(vBibTekst(TABLE_INVOICES, "#v048 #")) + Val(vBibTekst(TABLE_INVOICES, "#v047 #")) + Val(vBibTekst(TABLE_INVOICES, "#v046 #")) + Val(vBibTekst(TABLE_INVOICES, "#v049 #"))
        GridText = Dec$(dTTwb, MASK_SY(0)) + vbTab
        Intrastat.Show 1
    End If
    SS99 Right(Document_Key, 5), Ar   's001 of s003
    Document_Key = SleutelDok(Ar)
    DirekteAankoop.Caption = "Direkte aankoopverrichting         (" + Document_Key + ")"
    SchoonVegen_Click
    SSTab1.SetFocus
End If

End Sub

Private Sub Afsluiten_LostFocus()

Kontrole.Visible = True
Afsluiten.Visible = False

End Sub

Private Sub Annuleren_Click()

If AankoopDetail.ListCount = 0 Then
Else
    Msg = "Aanwezige bewerkingen negeren !  Bent U zeker ?"
    Ktrl = MsgBox(Msg, 292, "Aankoopverrichtingen sluiten")
    If Ktrl <> 6 Then
        Exit Sub
    End If
End If
Unload Me

End Sub

Private Sub ButtonOptimize_Click()

Dim HuidigeRekening As String
Dim TempoS As String
Dim TT As Integer
Dim TTT As Integer
Dim TTTT As Integer
Dim TTTTT As Integer
Dim twice As Integer
Dim totalAfterTwice As Integer

    If AankoopDetail.ListCount = 0 Then Exit Sub

    Msg = "Buiten producten, meermaals voorkomende lijnen samenvoegen tot ��n." + vbCrLf + vbCrLf
    Msg = Msg + "Bent U zeker ?"
    KtrlBox = MsgBox(Msg, 292)
    If KtrlBox = 6 Then
        For twice = 1 To 2
            TT = 0
            Do While TT < AankoopDetail.ListCount
                'TODO: check length of the line should be no more than 62
                If Len(AankoopDetail.List(TT)) = 62 Then
                    HuidigeRekening = Mid(AankoopDetail.List(TT), 1, 7)
                    TTTTT = TT
                    TTTT = TT + 1
                    Do While TTTT < AankoopDetail.ListCount
                        If HuidigeRekening = Mid(AankoopDetail.List(TTTT), 1, 7) And Len(AankoopDetail.List(TTTT)) = 62 Then
                            TempoS = AankoopDetail.List(TTTTT)
                            Mid(TempoS, 50, 12) = Dec$(Val(Mid(AankoopDetail.List(TTTT), 50, 12)) + Val(Mid(AankoopDetail.List(TTTTT), 50, 12)), "#########.00")
                            AankoopDetail.RemoveItem TTTT
                            AankoopDetail.List(TTTTT) = TempoS
                            TTT = TTT + 1
                        End If
                        TTTT = TTTT + 1
                    Loop
                End If
                TT = TT + 1
            Loop
            totalAfterTwice = totalAfterTwice + TTT
        Next
        MsgBox "Einde optimalisatie.  " + Format(totalAfterTwice) + " lijnen samengevoegd."
        'RefreshBTW
    End If

End Sub

Private Sub cbAfbeelding_Click()

    Dim flHier As Integer
    Dim BestandHier As String
    
    'Dim miDoc As MODIctl.Document
    'Set miDoc = New MODIctl.Document
    
    If LaadTekst("dnnInstellingen", "Mario") = "" Then
        MsgBox "Nieuwe PC of nog geen instellingen voor PDF Postvak In.  Aanbevolen in te te stellen a.u.b. via submenu DotNetNuke.", vbInformation
        Mim.Teken.InitDir = LOCATION_COMPANYDATA
    Else
        Mim.Teken.InitDir = LaadTekst("dnnInstellingen", "Mario")
    End If
    Mim.Teken.Filter = "Acrobat bestanden (*.pdf)|*.pdf"
    Mim.Teken.filename = ""
    Mim.Teken.ShowOpen
    BestandHier = LCase(Mim.Teken.filename)
    If BestandHier = "" Then
        Exit Sub
    Else
        If Right(BestandHier, 3) <> "pdf" Then
            MsgBox "Enkel *.pdf bestanden zijn toegelaten", vbExclamation
            Exit Sub
        End If
    End If
    GeScanBestand = Mim.Teken.filename
    Screen.MousePointer = vbHourglass
    On Local Error Resume Next
    OLE1.CreateEmbed GeScanBestand
    Screen.MousePointer = vbNormal
    
End Sub



Private Sub cbCheckTools_Click()

    Load ValidatingTool
    ValidatingTool.Caption = "Leveranciers Peppol Tools"
    ValidatingTool.tbCompanyNumber = supplierVatNumber
    ValidatingTool.tbVatNumber = supplierCountryCode + supplierVatNumber
    ValidatingTool.tbPeppolID = "0208:" + supplierVatNumber
    ValidatingTool.Show 1
    
    'If Trim(ValidatingTool.TextBoxSupportedDocuments) = "" Then
    'Else
    '    If rsSupplier("v410") = ValidatingTool.TextBoxSupportedDocuments.Text Then
    '        'MsgBox "Er zijn geen wijzigingen aan ondersteunde documenten", vbInformation
    '    Else
    '        Msg = "Ondersteunde Peppol verrichtingen bijwerken voor deze leverancier" & vbCrLf & vbCrLf
    '        Msg = Msg + "Bent U zeker"
    '        KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2 + vbQuestion)
    '        If KtrlBox = vbYes Then
    '            rsSupplier("v410") = ValidatingTool.TextBoxSupportedDocuments.Text
    '            rsSupplier.Update
                
    '            MsgBox "Ondersteunde verrichtingen bijgewerkt. Herstart de aankoopprocedure", vbInformation
    '            Schoon
    '        End If
    '    End If
    'End If
    Unload ValidatingTool

End Sub

Private Sub cbImportUBL_Click()

    Dim counterToBook As Integer
    Dim counterBooked As Integer

    Dim flHier As Integer
    Dim ublFileUrl As String
    
    Screen.MousePointer = vbHourglass
    On Local Error Resume Next
    
    FormB2BAMonitor.Hide
        
    FormB2BAMonitor.mfgToBook.Clear
    FormB2BAMonitor.mfgToBook.Cols = 6
    FormB2BAMonitor.mfgToBook.Rows = 2
                        
    FormB2BAMonitor.mfgToBook.TextMatrix(0, 0) = "Leverancier"
    FormB2BAMonitor.mfgToBook.TextMatrix(0, 1) = "Ondernemingsnummer"
    FormB2BAMonitor.mfgToBook.TextMatrix(0, 2) = "Type Doc."
    FormB2BAMonitor.mfgToBook.TextMatrix(0, 3) = "Datum Doc"
    FormB2BAMonitor.mfgToBook.TextMatrix(0, 4) = "Vervaldag"
    FormB2BAMonitor.mfgToBook.TextMatrix(0, 5) = "Bestand"
    
    FormB2BAMonitor.mfgToBook.ColWidth(0) = 2130
    FormB2BAMonitor.mfgToBook.ColWidth(1) = 1710
    FormB2BAMonitor.mfgToBook.ColWidth(2) = 1110
    FormB2BAMonitor.mfgToBook.ColWidth(3) = 990
    FormB2BAMonitor.mfgToBook.ColWidth(4) = 990
    FormB2BAMonitor.mfgToBook.ColWidth(5) = 2130
    
    FormB2BAMonitor.mfgBooked.Clear
    FormB2BAMonitor.mfgBooked.Cols = 6
    FormB2BAMonitor.mfgBooked.Rows = 2
                        
    FormB2BAMonitor.mfgBooked.TextMatrix(0, 0) = "Leverancier"
    FormB2BAMonitor.mfgBooked.TextMatrix(0, 1) = "Nummer Doc."
    FormB2BAMonitor.mfgBooked.TextMatrix(0, 2) = "Datum Doc"
    FormB2BAMonitor.mfgBooked.TextMatrix(0, 3) = "Vervaldag"
    FormB2BAMonitor.mfgBooked.TextMatrix(0, 4) = "Status"
    FormB2BAMonitor.mfgBooked.TextMatrix(0, 5) = "Bestand"
    
    FormB2BAMonitor.mfgBooked.ColWidth(0) = 2130
    FormB2BAMonitor.mfgBooked.ColWidth(1) = 1110
    FormB2BAMonitor.mfgBooked.ColWidth(2) = 990
    FormB2BAMonitor.mfgBooked.ColWidth(3) = 990
    FormB2BAMonitor.mfgBooked.ColWidth(4) = 1710
    FormB2BAMonitor.mfgBooked.ColWidth(5) = 2130
    
    Dim c As Long
    For c = 0 To 5
        FormB2BAMonitor.mfgToBook.ColAlignment(c) = flexAlignLeftCenter
        FormB2BAMonitor.mfgBooked.ColAlignment(c) = flexAlignLeftCenter
    Next
    
    Mim.Teken.InitDir = LOCATION_COMPANYDATA & "peppol\in\"
    ifSupplierInsertWarning = False
    
    Dim sPath As String
    Dim sFile As String
    
    sPath = LOCATION_COMPANYDATA & "peppol\in\"
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    sFile = Dir$(sPath & "*.xml")          ' first match
    Do While sFile <> ""
        On Local Error Resume Next
        Call ReadUblDocument(sPath & sFile, False)
        
        If CheckSupplier Then
            'For T = 0 To UBound(uitwisselingOMSArray): Print T; " "; uitwisselingOMSArray(T); ": "; uitwisselingDATAArray(T): Next
            '00 UBL VersionID: 2.1
            '01 documentIdToCheck:   V0202550406 -220935
            '02 dateSellerDocumentToCheck: 2025-11-08
            '03 dateExpiringDocumentToCheck: 2025-11-08
            '04 invoiceTypeCode: 380
            '05 supplierCompanyIdToCheck: 0440058217
            '06 supplierID:
            '07 supplierNameToCheck: Bv Zakenkantoor Hedwig Roelandt En Jos Vermoesen
            '08 supplierStreetToCheck: Grote Baan 141
            '09 supplierCityToCheck:           Herdersem
            '10 supplierPostalCodeToCheck: 9310
            '11 supplierCountryCodeToCheck:   BE
            '12 supplierVatNumberToCheck:   BE0440058217
            '13 paymentMeansCode: 1
            '14 payReferenceToCheck:   V0202550406 -220935
            '15 supplierIBANToCheck:   BE83891854037015
            '16 paySupplierNameToCheck:
            '17 supplierBICToCheck:   VDSPBE91
            '18 supplierDomMandate:
            '19 supplierDomDate:
            '20  totalExclusiveVAT: 110.37
            '21  totalInclusiveVAT: 133.55
            '22 customerAccountID:
            '23  customerEndpointID: 0874419267
            '24  customerName: Vermoesen Beheer CommV
            '25  customerStreetName: Grote Baan 141
            '26 customerCityName:   Herdersem
            '27  customerPostalZone: 9310
            '28 customerCountryCode:   BE
            '29 customerTaxID:           BE0874419267
            '30 customerTaxScheme:   VAT
            
            Dim entry As String
                            
            Dim documentCode As String
            Dim listItemToAdd As String
            
            Dim result As String
            result = CheckDocument
            Select Case uitwisselingDATAArray(4)
                Case "071", "084"
                    documentCode = "Debetnota" ' (" + uitwisselingDATAArray(4) + ")"
                Case "380", "386"
                    documentCode = "Factuur" ' (" + uitwisselingDATAArray(4) + ")"
                Case "381"
                    documentCode = "Creditnota" ' (" + uitwisselingDATAArray(4) + ")"
                Case "575"
                    documentCode = "Verzekering" ' (" + uitwisselingDATAArray(4) + ")"
                Case Else
                    documentCode = uitwisselingDATAArray(4)
            
            End Select
            
            
            Dim dateDocument As String
            dateDocument = DATE_TEXT(Replace(uitwisselingDATAArray(2), "-", ""))
            
            Dim dateExpiringDocument As String
            If uitwisselingDATAArray(4) = "380" Then
                dateExpiringDocument = DATE_TEXT(Replace(uitwisselingDATAArray(3), "-", ""))
            Else
                dateExpiringDocument = dateDocument
            End If
            
            listItemToAdd = uitwisselingDATAArray(7) & " - " & documentCode & " - " & dateDocument & " vvd. " & dateExpiringDocument
            
            If Mid(result, 1, 9) = "Ingeboekt" Then
                counterBooked = counterBooked + 1
                entry = uitwisselingDATAArray(7) & vbTab & XLogKey & vbTab & dateDocument & vbTab & dateExpiringDocument & vbTab & Mid(result, 10) & vbTab & sFile
                FormB2BAMonitor.mfgBooked.AddItem entry, counterBooked
            Else
                counterToBook = counterToBook + 1
                entry = uitwisselingDATAArray(7) & vbTab & uitwisselingDATAArray(5) & vbTab & documentCode & vbTab & dateDocument & vbTab & dateExpiringDocument & vbTab & sFile
                FormB2BAMonitor.mfgToBook.AddItem entry, counterToBook
            End If
        Else
            'TODO: supplier not found
            'FormB2BMonitor.ListBoxToBook.AddItem sFile & " {" & uitwisselingDATAArray(5) & listItemToAdd & ", Leverancier niet gevonden/bij te werken}"
        End If
        sFile = Dir$()                      ' subsequent matches
    Loop
    Screen.MousePointer = vbNormal
    XLogKey = ""
    FormB2BAMonitor.mfgToBook.SetFocus
    
    If Not purchasePeppolTODOShowed Then
        purchasePeppolTODOShowed = True
        If FormB2BAMonitor.mfgToBook.Rows = 2 Then
            'no documents to book
            documentLinesDATA = ""
            Unload FormB2BAMonitor
            Exit Sub
        Else
            'documents to book message until 28/02/2026
            'If #2/28/2026# > Date Then
            '    Msg = Trim(Str(FormB2BAMonitor.mfgToBook.Rows - 2)) & " in te boeken Peppol aankoopdocumenten." & vbCrLf
            '    Msg = Msg & "Boek ze spoedig in via [Document Inladen]." & vbCrLf & vbCrLf
            '    Msg = Msg & "Manueel ingeboekte Peppol aankoopdocumenten zijn niet controleerbaar." & vbCrLf & vbCrLf
            '    Msg = Msg & "Dit bericht verschijnt t.e.m. 28/02/2026"
            '    MsgBox Msg, vbInformation, "Documenten in te boeken"
            'End If
        End If
        
        'TODO message when documents have to be booked
        FormB2BAMonitor.Show 1
    Else
        FormB2BAMonitor.Show 1
    End If
        
    ifSupplierInsertWarning = True
    ublFileUrl = LCase(XLogKey)
    If ublFileUrl = "" Then
        ifSupplierInsertWarning = True
        Exit Sub
    ElseIf Dir(ublFileUrl) = "" Then
        Exit Sub
    Else
        If Right(ublFileUrl, 3) <> "xml" Then
            MsgBox "Enkel *.xml bestanden zijn toegelaten", vbExclamation
            Exit Sub
        End If
    End If
        
    Screen.MousePointer = vbHourglass
    On Local Error Resume Next
    
    Call ReadUblDocument(ublFileUrl, False)
    
    ReDim uitwisselingOMSArray(50)
    ReDim uitwisselingDATAArray(50)
    uitwisselingOMSArray() = Split(uitwisselingOMS, vbTab)
    uitwisselingDATAArray() = Split(uitwisselingDATA, vbTab)
            
    Screen.MousePointer = vbNormal
    'Supplier lookup with companyId
    If CheckSupplier Then
        'Found! Now search if not already booked
        If CheckDocument() Then
            'KtrlBox = MsgBox("Een ingesloten PDF document uitpakken.", vbQuestion + vbYesNo + vbDefaultButton1)
            'If KtrlBox = vbYes Then
            '    Dim r As Long
            '    ExtractPdfFromUBLDocument (ublFileUrl)
        
            '    Dim fileUrlAsPdf As String
            '    fileUrlAsPdf = Left(ublFileUrl, Len(ublFileUrl) - 3) + "pdf"
            '    r = ShellExecute(0, "open", fileUrlAsPdf, 0, 0, 1)
            'End If
    
            bGet TABLE_SUPPLIERS, 0, rsSupplier("A110")
            If Ktrl Then
                MsgBox "stop"
            Else
                RecordToVeld TABLE_SUPPLIERS
                InstalSupplierForUBL
            End If
        End If
    End If
    
End Sub


Private Sub cmdSQLInfo_Click()

bGet TABLE_SUPPLIERS, 0, Left(LeverancierInfo.Caption, 12)
If Ktrl Then
    MsgBox "Onlogica."
    cmdSQLInfo.Visible = False
Else
    RecordToVeld TABLE_SUPPLIERS
    If vBibTekst(TABLE_SUPPLIERS, "#v254 #") <> "" Then
        KtrlBox = SQLPopUp(vBibTekst(TABLE_SUPPLIERS, "#v254 #"), bstNaam(TABLE_SUPPLIERS), "A110", vBibTekst(TABLE_SUPPLIERS, "#A110 #"))
    Else
        cmdSQLInfo.Visible = False
    End If
End If

End Sub


Sub VernieuwDirecteAankoopEUR()

'    MsgBox "Stop"
    
End Sub

Private Sub cmdXLog_Click()

bGet TABLE_SUPPLIERS, 0, supplierA110
If Ktrl = 0 Then
    RecordToVeld TABLE_SUPPLIERS
End If
           
If TeleBibClick(TABLE_SUPPLIERS) = False Then
    SSTab1.Tab = 0
Else
    Msg = "Gegevens bestaande '" + bstNaam(TABLE_SUPPLIERS) + "'-fiche wijzigen.  Bent U zeker ?"
    KtrlBox = MsgBox(Msg, 292)
    If KtrlBox = 6 Then
        bUpdate TABLE_SUPPLIERS, 0
    End If
    SSTab1.Tab = 0
    InstalLeverancier
End If

End Sub


Private Sub Form_Load()

purchasePeppolTODOShowed = False

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

If String99(READING, 100) = "1" Then
    documentRefIsOGM = True
Else
    documentRefIsOGM = False
End If

If Trim(String99(READING, 20)) = "7" Then
    adminNoVat = True
Else
    adminNoVat = False
End If

If (Right(LOCATION_COMPANYDATA, 5) = "\098\" Or Right(LOCATION_COMPANYDATA, 5) = "\099\") Then
    Me.TextBoxWarningTestCompany.Visible = True
End If

SSTab1.Tab = 1
Me.cbImportUBL.Enabled = True
'rsMAR(TABLE_INVOICES).Close
Ktrl = ktrlBLOBRecord
If Ktrl = False Then
    AutoUnloadBedrijf
    Exit Sub
End If

'If bhEuro Then
'    cmdSwitch.Caption = "Ingave in EUR"
'Else
'    cmdSwitch.Caption = "Ingave in BEF"
'End If

Dim Tel As Integer

' Add ColumnHeaders. The width of the columns is
' the width of the control divided by the number of
' ColumnHeader objects.
ListView1.ColumnHeaders.Add , , "Naam", 3270
ListView1.ColumnHeaders.Add , , "IdCode", 930
          
' Set View property to Report.
ListView1.View = lvwReport
        
Screen.MousePointer = vbHourglass
Top = 0
Left = 0

For Tel = 16 To 19
    rbtwVAK(Tel - 16) = String99(READING, Tel)
    rbtwVAK(Tel - 12) = String99(READING, Tel + 6)
Next

LeverancierRekening = String99(READING, 10)
sIsIntraFlg = String99(READING, 200)


PriveRekening = String99(READING, 145)

GrensDetail(0) = ""
FL99_RECORD = String99(READING, 148)
    Mid(GrensDetail(0), 1, 7) = Left(FL99_RECORD, 7)
FL99_RECORD = String99(READING, 149)
    Mid(GrensDetail(0), 8, 7) = Left(FL99_RECORD, 7)

GrensDetail(1) = ""
FL99_RECORD = String99(READING, 146)
    Mid(GrensDetail(1), 1, 7) = Left(FL99_RECORD, 7)
FL99_RECORD = String99(READING, 147)
    Mid(GrensDetail(1), 8, 7) = Left(FL99_RECORD, 7)

GrensDetail(2) = ""
FL99_RECORD = String99(READING, 150)
    Mid(GrensDetail(2), 1, 7) = Left(FL99_RECORD, 7)
FL99_RECORD = String99(READING, 151)
    Mid(GrensDetail(2), 8, 7) = Left(FL99_RECORD, 7)

GrensDetail(3) = ""
FL99_RECORD = String99(READING, 152)
    Mid(GrensDetail(3), 1, 7) = Left(FL99_RECORD, 7)
FL99_RECORD = String99(READING, 153)
    Mid(GrensDetail(3), 8, 7) = Left(FL99_RECORD, 7)

'InstalleerRecenteCrediteuren
Msg = LaadTekst("DirekteAankoop", "startBlad")
If Msg = "" Then Msg = "True"
StartBlad = Msg
Schoon

Screen.MousePointer = vbNormal
If Not purchasePeppolTODOShowed Then
    cbImportUBL_Click
End If

End Sub

Private Sub Form_Resize()

    AankoopDetail.Width = DirekteAankoop.Width - 250
    Me.SSTab1.SetFocus
    If ifSupplierInsertWarning Then
        BasisB(TABLE_SUPPLIERS).SetFocus
    End If
    
End Sub


Private Sub InstalLeverancier()
Dim Lever As String
Dim T       As Integer
Dim X       As Integer

Me.cbImportUBL.Enabled = False

AankoopDetail.Enabled = True
'InvoerLijn.Enabled = True

supplierA110 = Trim(vBibTekst(TABLE_SUPPLIERS, "#A110 #"))

Lever = vbCrLf + vBibTekst(TABLE_SUPPLIERS, "#A100 #") + vbCrLf + vBibTekst(TABLE_SUPPLIERS, "#A125 #") + vbCrLf + vBibTekst(TABLE_SUPPLIERS, "#A104 #") + vbCrLf + vBibTekst(TABLE_SUPPLIERS, "#A109 #") + " " + vBibTekst(TABLE_SUPPLIERS, "#A107 #") + " " + vBibTekst(TABLE_SUPPLIERS, "#A108 #")
Schoonvegen.Enabled = True
'ReturnRooster.Enabled = True

If vBibTekst(TABLE_SUPPLIERS, "#v149 #") = "" Then
    MsgBox "Landnummer is verplicht !"
    Exit Sub
ElseIf vBibTekst(TABLE_SUPPLIERS, "#v149 #") = "002" Then
    LeverancierInfo.Caption = vSet(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 12) + "* Binnenland * " + Lever
    AankoopFlg = 0
    Medekontraktant.Enabled = True
    TekstInfo7.Visible = False
    TekstInfo7.text = TekstInfo5.text
    Label1(12).Visible = False
ElseIf InStr(SISO, vBibTekst(TABLE_SUPPLIERS, "#v149 #")) Then
    LeverancierInfo.Caption = vSet(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 12) + "* E.U. Verschuldigde BTW * " + Lever
    AankoopFlg = 1
    Medekontraktant.Enabled = False
    TekstInfo7.Visible = True
    TekstInfo7.text = "0"
    Label1(12).Visible = True
Else
    LeverancierInfo.Caption = vSet(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 12) + "* Niet E.U. + BTW ! *" + Lever
    AankoopFlg = 2
    Medekontraktant.Enabled = False
    TekstInfo7.Visible = True
    TekstInfo7.text = "0"
    Label1(12).Visible = True
End If

If Trim(vBibTekst(TABLE_SUPPLIERS, "#A161 #")) <> "" Then
    supplierCompanyId = Trim(vBibTekst(TABLE_SUPPLIERS, "#v404 #"))
    supplierCountryCode = Trim(vBibTekst(TABLE_SUPPLIERS, "#v150 #"))
    supplierVatNumber = Trim(vBibTekst(TABLE_SUPPLIERS, "#A161 #"))
    
End If

TekstInfo0.Enabled = True
TekstInfo1.Enabled = True
TekstInfo2.Enabled = True
TekstInfo3.Enabled = True
TekstInfo4.Enabled = True
TekstInfo12.Enabled = True
TekstInfo5.Enabled = True
TekstInfo6.Enabled = True
TekstInfo7.Enabled = True

sMuntLever = UCase(vBibTekst(TABLE_SUPPLIERS, "#vs03 #"))
bGet TABLE_VARIOUS, 1, vSet("10" + sMuntLever, 20)
If Ktrl Then
    MsgBox "Dagkoers voor muntkode " & sMuntLever & " niet te vinden !  Eerst aanmaken via gebruikersfiches a.u.b."
    If bhEuro Then
        sMuntLever = "EUR"
    Else
        sMuntLever = "BEF"
    End If
    TekstInfo9.text = Dec$(1, "###.########")
Else
    RecordToVeld TABLE_VARIOUS
    TekstInfo9.text = Dec$(Val(vBibTekst(TABLE_VARIOUS, "#v040 #")), "###.########")
End If

TekstInfo9.text = Dec$(1, "##0.########")
dMuntL = Val(TekstInfo9.text)

If vBibTekst(TABLE_SUPPLIERS, "#v151 #") = "1" Then
    Medekontraktant.Value = 1
Else
    Medekontraktant.Value = 0
End If
If vBibTekst(TABLE_SUPPLIERS, "#v163 #") = "1" Then
    StockBeheer.Value = 1
Else
    StockBeheer.Value = 0
End If

Err = 0
On Local Error Resume Next

If vSet(vBibTekst(TABLE_SUPPLIERS, "#v016 #"), 7) <> Space$(7) Then
    If AankoopDetail.ListCount = 0 Then
    RasterSchoon
    VeldRekening = vBibTekst(TABLE_SUPPLIERS, "#v016 #")
    VeldNaam = ""
    VeldBedrag = ""
    GridText = VeldRekening + Chr$(124) + VeldNaam + Chr$(124) + VeldBedrag + Chr$(124)
    Wijzigen.Show 1
    If GridText = "" Then
    Else
        If RekeningOK(Left(GridText, 7)) Then
            AankoopDetail.AddItem GridText
            X = InvestKtrl()
        End If
    End If
    End If
End If

If Left(vBibTekst(TABLE_SUPPLIERS, "#v162 #"), 3) = "440" Then
    TekstInfo3.text = vBibTekst(TABLE_SUPPLIERS, "#v162 #")
    bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_SUPPLIERS, "#v162 #"), 7)
    If Ktrl Then
        TekstInfo3.text = LeverancierRekening
    End If
Else
    TekstInfo3.text = LeverancierRekening
End If
TekstInfo10.Enabled = True

TekstInfo2.text = VValdag((TekstInfo0.text), vBibTekst(TABLE_SUPPLIERS, "#vs04 #"))
If vBibTekst(TABLE_SUPPLIERS, "#v017 #") = "1" Then
    TekstInfo4.AutoTab = True
    TekstInfo4.ClipMode = 1
    TekstInfo4.PromptInclude = False
    TekstInfo4.Mask = "+++###/####/#####+++"
Else
    TekstInfo4.AutoTab = False
    TekstInfo4.ClipMode = 0
    TekstInfo4.PromptInclude = True
    TekstInfo4.Mask = ""
End If
TekstInfo4.text = vBibTekst(TABLE_SUPPLIERS, "#v412 #")
If Trim(TekstInfo4.text) = "" Then
Else
    TekstInfo4.Enabled = False
End If

If vBibTekst(TABLE_SUPPLIERS, "#v254 #") <> "" Then
    cmdSQLInfo.Visible = True
End If
Me.cmdXLog.Enabled = True
Me.LabelInfoXlog.Visible = False
SSTab1.Tab = 0

If supplierCompanyId = "" Then
ElseIf supplierCountryCode = "BE" Then
    Dim beSupplierPeppolCheck As Boolean
    beSupplierPeppolCheck = CheckSupplierDocuments(supplierA110)
    If beSupplierPeppolCheck = False Then
        MsgBox "Deze BE Leverancier dient via Peppol te werken" & vbCrLf & vbCrLf & "In principe mag je geen aankoopdocumenten MANUEEL inbrengen", vbExclamation
    End If
End If

Kontrole.Enabled = True
DirekteAankoop.AankoopDetail.SetFocus

End Sub

Private Sub InstalSupplierForUBL()
Dim Lever As String
Dim T       As Integer
Dim X       As Integer

Me.cbImportUBL.Enabled = False
AankoopDetail.Enabled = True
'InvoerLijn.Enabled = True

supplierA110 = Trim(vBibTekst(TABLE_SUPPLIERS, "#A110 #"))

Lever = vbCrLf + vBibTekst(TABLE_SUPPLIERS, "#A100 #") + vbCrLf + vBibTekst(TABLE_SUPPLIERS, "#A125 #") + vbCrLf + vBibTekst(TABLE_SUPPLIERS, "#A104 #") + vbCrLf + vBibTekst(TABLE_SUPPLIERS, "#A109 #") + " " + vBibTekst(TABLE_SUPPLIERS, "#A107 #") + " " + vBibTekst(TABLE_SUPPLIERS, "#A108 #")
Schoonvegen.Enabled = True
'ReturnRooster.Enabled = True

If vBibTekst(TABLE_SUPPLIERS, "#v149 #") = "" Then
    MsgBox "Landnummer is verplicht !"
    Exit Sub
ElseIf vBibTekst(TABLE_SUPPLIERS, "#v149 #") = "002" Then
    LeverancierInfo.Caption = vSet(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 12) + "* Binnenland * " + Lever
    AankoopFlg = 0
    Medekontraktant.Enabled = True
    TekstInfo7.Visible = False 'btw verschuldigd
    TekstInfo7.text = TekstInfo5.text 'btw aftrekbaar
    Label1(12).Visible = False
ElseIf InStr(SISO, vBibTekst(TABLE_SUPPLIERS, "#v149 #")) Then
    LeverancierInfo.Caption = vSet(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 12) + "* E.U. Verschuldigde BTW * " + Lever
    AankoopFlg = 1
    Medekontraktant.Enabled = False
    TekstInfo7.Visible = True
    TekstInfo7.text = "0"
    Label1(12).Visible = True
Else
    LeverancierInfo.Caption = vSet(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 12) + "* Niet E.U. + BTW ! *" + Lever
    AankoopFlg = 2
    Medekontraktant.Enabled = False
    TekstInfo7.Visible = True
    TekstInfo7.text = "0"
    Label1(12).Visible = True
End If

If Trim(vBibTekst(TABLE_SUPPLIERS, "#A161 #")) <> "" Then 'VAT number

    supplierCompanyId = Trim(vBibTekst(TABLE_SUPPLIERS, "#v404 #"))
    supplierCountryCode = Trim(vBibTekst(TABLE_SUPPLIERS, "#v150 #"))
    supplierVatNumber = Trim(vBibTekst(TABLE_SUPPLIERS, "#A161 #"))
    
    'If supplierCompanyId = supplierCompanyIdToCheck Then
    '    If supplierCountryCode = supplierCountryCodeToCheck Then
    '        'ok
    '    Else
    '        MsgBox "Stop"
    '    End If
    'Else
    '    MsgBox "Stop"
    'End If
            
    'checken met UBL document zou identiek moeten zijn
    
End If

TekstInfo0.Enabled = True
TekstInfo1.Enabled = True
TekstInfo2.Enabled = True
TekstInfo3.Enabled = True
TekstInfo4.Enabled = True
TekstInfo12.Enabled = True
TekstInfo5.Enabled = True
TekstInfo6.Enabled = True
TekstInfo7.Enabled = True

sMuntLever = UCase(vBibTekst(TABLE_SUPPLIERS, "#vs03 #"))
bGet TABLE_VARIOUS, 1, vSet("10" + sMuntLever, 20)
If Ktrl Then
    MsgBox "Dagkoers voor muntkode " & sMuntLever & " niet te vinden !  Eerst aanmaken via gebruikersfiches a.u.b."
    If bhEuro Then
        sMuntLever = "EUR"
    Else
        sMuntLever = "BEF"
    End If
    TekstInfo9.text = Dec$(1, "###.########")
Else
    RecordToVeld TABLE_VARIOUS
    TekstInfo9.text = Dec$(Val(vBibTekst(TABLE_VARIOUS, "#v040 #")), "###.########")
End If

If bhEuro Then
    If sMuntLever = "BEF" Then
        'cmdSwitch.Caption = "Ingave in BEF"
        'cmdSwitch.Enabled = True
        TekstInfo9.text = Dec$(1 / EURO, "##0.########")
    ElseIf sMuntLever = "EUR" Then
        'cmdSwitch.Caption = "Ingave in EUR"
        'cmdSwitch.Enabled = True
        TekstInfo9.text = Dec$(1, "##0.########")
    Else
        'cmdSwitch.Caption = "Ingave in EUR"
        'cmdSwitch.Enabled = False
        TekstInfo9.text = Dec$(1, "##0.########")
    End If
Else
    If sMuntLever = "EUR" Then
        'cmdSwitch.Caption = "Ingave in EUR"
        'cmdSwitch.Enabled = True
        TekstInfo9.text = Dec$(1 / EURO, "##0.########")
    ElseIf sMuntLever = "BEF" Then
        'cmdSwitch.Caption = "Ingave in BEF"
        'cmdSwitch.Enabled = True
        TekstInfo9.text = Dec$(1, "##0.########")
    Else
        'cmdSwitch.Caption = "Ingave in BEF"
        'cmdSwitch.Enabled = False
        'TekstInfo9.Text = Dec$(1, "##0.########")
    End If
End If
dMuntL = Val(TekstInfo9.text)

If vBibTekst(TABLE_SUPPLIERS, "#v151 #") = "1" Then
    Medekontraktant.Value = 1
Else
    Medekontraktant.Value = 0
End If
If vBibTekst(TABLE_SUPPLIERS, "#v163 #") = "1" Then
    StockBeheer.Value = 1
Else
    StockBeheer.Value = 0
End If

Err = 0
On Local Error Resume Next

If Left(vBibTekst(TABLE_SUPPLIERS, "#v162 #"), 3) = "440" Then
    TekstInfo3.text = vBibTekst(TABLE_SUPPLIERS, "#v162 #")
    bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_SUPPLIERS, "#v162 #"), 7)
    If Ktrl Then
        TekstInfo3.text = LeverancierRekening
    End If
Else
    TekstInfo3.text = LeverancierRekening
End If
TekstInfo10.Enabled = True

If vBibTekst(TABLE_SUPPLIERS, "#v017 #") = "1" Then
    TekstInfo4.AutoTab = True
    TekstInfo4.ClipMode = 1
    TekstInfo4.PromptInclude = False
    TekstInfo4.Mask = "+++###/####/#####+++"
Else
    TekstInfo4.AutoTab = False
    TekstInfo4.ClipMode = 0
    TekstInfo4.PromptInclude = True
    TekstInfo4.Mask = ""
    TekstInfo4.text = ""
End If
If vBibTekst(TABLE_SUPPLIERS, "#v254 #") <> "" Then
    cmdSQLInfo.Visible = True
End If
Me.cmdXLog.Enabled = True
Me.LabelInfoXlog.Visible = False
SSTab1.Tab = 0

TekstInfo1.text = DATE_TEXT(Replace(uitwisselingDATAArray(2), "-", "")) 'date document
TekstInfo0.text = TekstInfo1.text
TekstInfo2.text = DATE_TEXT(Replace(uitwisselingDATAArray(3), "-", "")) 'expiring
    
TekstInfo4.text = uitwisselingDATAArray(14)
TekstInfo12.text = uitwisselingDATAArray(1)

If adminNoVat Then
    TekstInfo5.text = "0"
Else
    If AankoopFlg = 0 Then
        TekstInfo5.text = Str((Val(uitwisselingDATAArray(21)) - Val(uitwisselingDATAArray(20))))
        TekstInfo5.Enabled = True
    Else
        MsgBox "EU aankoopdocument. Btw automatisch op 21%. Verbeter indien nodig.", vbExclamation
        TekstInfo5.text = Dec(Val(uitwisselingDATAArray(20)) * 21 / 100, ".##")
        TekstInfo7.text = TekstInfo5.text
        TekstInfo5.Enabled = True
        TekstInfo7.Enabled = False
    End If
End If
TekstInfo6.text = uitwisselingDATAArray(21)
TekstInfo1.Enabled = False
TekstInfo2.Enabled = False
TekstInfo4.Enabled = False
TekstInfo12.Enabled = False
'TekstInfo5.Enabled = False
TekstInfo6.Enabled = False
TekstInfo12.Enabled = False

    ReDim documentLinesOMSArray(20)
    documentLinesOMSArray() = Split(documentLinesOMS, vbTab)
    ParseToArray (documentLinesDATA)
    '? documentLinesOMS
    
    Dim i As Integer, j As Integer
    Dim squareCtrl As Currency
    Dim checkSquare As Currency
    Dim lineValue As Currency
    Dim squareDifference As Currency
    Dim lineStringValue As String

    squareCtrl = 0
    For i = LBound(documentLinesDATAArray, 1) + 1 To UBound(documentLinesDATAArray, 1) - 1
        If adminNoVat Then
            lineValue = Val(Dec((Val(documentLinesDATAArray(i, 7)) * (1 + Val(documentLinesDATAArray(i, 8)) / 100)), MASK_EURBH))
            squareCtrl = squareCtrl + lineValue
        Else
            squareCtrl = squareCtrl + Val(documentLinesDATAArray(i, 7))  ' (Val(documentLinesDATAArray(i, 5)) * Val(documentLinesDATAArray(i, 6)))
        End If
    Next i
    If adminNoVat Then
        checkSquare = Val(uitwisselingDATAArray(21))
    Else
        checkSquare = Val(uitwisselingDATAArray(20))
    End If
    
    squareDifference = checkSquare - squareCtrl
    If squareDifference = 0 Then
    Else
        If adminNoVat Then
            Msg = "Rekenverschil bij vierkantscontrole" & vbCrLf & vbCrLf
            Msg = Msg & "Controle der lijnen geeft: " & squareCtrl & vbCrLf
            Msg = Msg & "en totaal incl. BTW in UBL document geeft: " & Val(uitwisselingDATAArray(21)) & vbCrLf & vbCrLf
            Msg = Msg & "Administratie zonder BTW plicht" & vbCrLf
            Msg = Msg & "Verschil wordt verrekend in eerste lijn."
            MsgBox Msg, vbExclamation, "Vierkantscontrole der lijnen"
        Else
            Msg = "Rekenverschil bij vierkantscontrole" & vbCrLf & vbCrLf
            Msg = Msg & "Controle der lijnen geeft: " & squareCtrl & vbCrLf
            Msg = Msg & "en totaal excl. BTW in UBL document geeft: " & Val(uitwisselingDATAArray(20)) & vbCrLf & vbCrLf
            Msg = Msg & "Administratie met BTW plicht"
            Msg = Msg & "Verschil wordt verrekend in eerste lijn."
            MsgBox Msg, vbExclamation, "Vierkantscontrole der lijnen"
        End If
    End If
    
    If vSet(vBibTekst(TABLE_SUPPLIERS, "#v016 #"), 7) <> Space$(7) Then
        If AankoopDetail.ListCount = 0 Then
            RasterSchoon
            VeldRekening = vBibTekst(TABLE_SUPPLIERS, "#v016 #")
            For i = LBound(documentLinesDATAArray, 1) + 1 To UBound(documentLinesDATAArray, 1) - 1
                If adminNoVat Then
                    lineValue = Val(Dec(((Val(documentLinesDATAArray(i, 7)) * (1 + Val(documentLinesDATAArray(i, 8)) / 100))), MASK_EUR + "00"))
                Else
                    lineValue = Val(documentLinesDATAArray(i, 7))
                End If
                If squareDifference <> 0 Then
                    lineValue = lineValue + squareDifference
                    squareDifference = 0
                End If
                lineStringValue = Dec$((lineValue), MASK_EURBH)
                If documentLinesDATAArray(i, 3) <> "-" Then
                    VeldNaam = vSet(documentLinesDATAArray(i, 3), 40)
                Else
                    VeldNaam = vSet(documentLinesDATAArray(i, 4), 40)
                End If
                GridText = VeldRekening + Chr$(124) + VeldNaam + Chr$(124) + lineStringValue + Chr$(124)
                If RekeningOK(Left(GridText, 7)) Then
                    AankoopDetail.AddItem GridText
                    X = InvestKtrl()
                End If
            Next i
        End If
    Else
        If AankoopDetail.ListCount = 0 Then
            RasterSchoon
            VeldRekening = vSet("604000", 7)
            For i = LBound(documentLinesDATAArray, 1) + 1 To UBound(documentLinesDATAArray, 1) - 1
                If adminNoVat Then
                    lineValue = Val(Dec(((Val(documentLinesDATAArray(i, 7)) * (1 + Val(documentLinesDATAArray(i, 8)) / 100))), MASK_EUR + "00"))
                Else
                    lineValue = Val(documentLinesDATAArray(i, 7))
                End If
                If squareDifference <> 0 Then
                    lineValue = lineValue + squareDifference
                    squareDifference = 0
                End If
                lineStringValue = Dec$((lineValue), MASK_EURBH)
                If documentLinesDATAArray(i, 3) = "-" Then
                    VeldNaam = vSet(documentLinesDATAArray(i, 4), 40)
                Else
                    VeldNaam = vSet(documentLinesDATAArray(i, 3), 40)
                End If
                GridText = VeldRekening + Chr$(124) + VeldNaam + Chr$(124) + lineStringValue + Chr$(124)
                If RekeningOK(Left(GridText, 7)) Then
                    AankoopDetail.AddItem GridText
                    X = InvestKtrl()
                End If
            Next i
        End If
    End If

If uitwisselingDATAArray(4) = "380" Then
    Me.AankoopOptie(0).Value = True
ElseIf uitwisselingDATAArray(4) = "381" Then
    Me.AankoopOptie(1).Value = True
End If

Kontrole.Enabled = True
DirekteAankoop.AankoopDetail.SetFocus

End Sub



Private Sub ParseToArray(inputText As String)
    
    Dim lines() As String
    Dim fields() As String
    Dim result() As String
    Dim i As Long, j As Long
    Dim maxCols As Long
    
    On Local Error GoTo 0
    
    ' Split into lines
    lines = Split(inputText, vbCrLf)
    
    ' Find maximum number of columns
    For i = LBound(lines) To UBound(lines)
        fields = Split(lines(i), vbTab)
        If UBound(fields) + 1 > maxCols Then
            maxCols = UBound(fields) + 1
        End If
    Next i
    
    ' Dimension result array: rows = number of lines, cols = max fields
    ReDim documentLinesDATAArray(UBound(lines), maxCols - 1)
    
    ' Fill the array
    For i = LBound(lines) To UBound(lines)
        fields = Split(lines(i), vbTab)
        For j = LBound(fields) To UBound(fields)
            documentLinesDATAArray(i, j) = fields(j)
        Next j
    Next i
    
End Sub

Private Function InvestKtrl() As Integer
Dim fRekNum As String * 7

fRekNum = Left(GridText, 7)
If fRekNum < Left(GrensDetail(0), 7) Or fRekNum > Right(GrensDetail(0), 7) Then
Else
    GridText = TekstInfo0.text + Dec$(Val(Mid(GridText, 50, 12)), MASK_EURBH)
    InvesteringsFiche.Show 1
End If

End Function

Private Sub Form_Unload(Cancel As Integer)

    'Me.MiDocView1.filename = ""
    
End Sub

Private Sub Kontrole_Click()
Dim T As Integer
Dim BedragTotaal As Currency
Dim BedragBtwAftrekbaas As Currency

Dim BedragTot As Currency
Dim BedragBtw5 As Currency
Dim BedragBtw6 As Currency
Dim BedragBtw7 As Currency
Dim BedragBtw8 As Currency

If Trim(TekstInfo12.text) = "" Then
    MsgBox "Referte Document is verplicht", vbInformation
    TekstInfo12.SetFocus
    Exit Sub
End If

If Trim(TekstInfo4.text) = "" Then
    'MsgBox "BetaalReferte is verplicht", vbInformation
    'TekstInfo4.SetFocus
    'Exit Sub
End If


If Not DATE_CHECK((TekstInfo0.text), PERIODAS_TEXT) Then
    BYPERDAT.WindowState = 0
    BYPERDAT.SetFocus
    BYPERDAT.PeriodeBoekjaar.SetFocus
    Exit Sub
End If

If DATE_KEY(TekstInfo1.text) < Left(BOOKYEAR_FROMTO, 8) Or DATE_KEY(TekstInfo1.text) > Right(BOOKYEAR_FROMTO, 8) Then
    Msg = "Datum aankoopdocument valt BUITEN het actieve boekjaar." & vbCr
    Msg = Msg & "De optie 'boekhoudcontrole' in balans leveranciers zal" & vbCr
    Msg = Msg & "mogelijk niet goed functioneren." & vbCr & vbCr
    Msg = Msg & "Controleer eventueel."
    MsgBox Msg, vbExclamation
End If

Kontrole.Default = True
For T = 0 To AankoopDetail.ListCount - 1
    BedragTot = Val(Mid(AankoopDetail.List(T), 50, 12))
    If bhEuro Then
        BedragTotaal = BedragTotaal + BedragTot
    Else
        BedragTotaal = BedragTotaal + BedragTot
    End If
Next
BedragBtw5 = Val(TekstInfo5.text)
BedragBtw6 = Val(TekstInfo6.text)
BedragBtw7 = Val(TekstInfo7.text)
BedragTotaal = BedragTotaal + BedragBtw5 - BedragBtw7 + BedragBtw8
If BedragTotaal = 0 Then Exit Sub
If BedragTotaal <> BedragBtw6 Then
    MsgBox "Totaal dokument " & BedragBtw6 & vbCr & "en volgens berekening is het " & BedragTotaal & vbCrLf & vbCrLf & "Kontroleer a.u.b.!", 0, "Totaalkontrole"
    If TekstInfo6.Enabled Then
        TekstInfo6.SetFocus
    End If
    Exit Sub
Else
    Afsluiten.Visible = True
    Afsluiten.Enabled = True
    Kontrole.Visible = False
    Afsluiten.SetFocus
End If

End Sub




Private Sub ListView1_Click()
    
    If ListView1.ListItems.Count Then
        lvLeverancier = ListView1.SelectedItem.SubItems(1)
        bGet TABLE_SUPPLIERS, 0, lvLeverancier
        If Ktrl Then
            'MsgBox "stop"
        Else
            RecordToVeld TABLE_SUPPLIERS
            InstalLeverancier
            SSTab1.Tab = 0:
        End If
    Else
        KtrlBox = MsgBox("Historiek van reeds bestaande leveranciers tonen", vbQuestion + vbYesNo + vbDefaultButton2)
        If KtrlBox = vbYes Then
            InstalleerRecenteCrediteuren
        End If
    End If
    
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 13 Then ListView1_Click

End Sub


Private Sub Medekontraktant_Click()

If Medekontraktant.Value = 1 Then
    TekstInfo7.Visible = True
    TekstInfo5.text = "0"
    TekstInfo6.text = "0"
    TekstInfo7.text = "0" 'TekstInfo(5).Text
    Label1(12).Visible = True
Else
    TekstInfo7.Visible = False
    TekstInfo5.text = "0"
    TekstInfo6.text = "0"
    TekstInfo7.text = "0"
    Label1(12).Visible = False
End If


End Sub

Private Sub RasterSchoon()

AankoopDetail.Clear

End Sub

Private Function RekeningOK(fRekeningNummer As String) As Integer
Dim T       As Integer
Dim TFlag   As Integer

TFlag = False
For T = 0 To 3
    If fRekeningNummer >= Left(GrensDetail(T), 7) And fRekeningNummer <= Right(GrensDetail(T), 7) Then
        TFlag = True
        Exit For
    End If
Next
If TFlag = False Then
          Msg = "Uw rekening : " + fRekeningNummer + " past niet in de begrenzing." + vbCrLf + vbCrLf
    Msg = Msg + "Investeringen  : vanaf " + Left(GrensDetail(0), 7) + " tot " + Right(GrensDetail(0), 7) + vbCrLf
    Msg = Msg + "Schulden/priv� : vanaf " + Left(GrensDetail(1), 7) + " tot " + Right(GrensDetail(1), 7) + vbCrLf
    Msg = Msg + "Handelsgoed    : vanaf " + Left(GrensDetail(2), 7) + " tot " + Right(GrensDetail(2), 7) + vbCrLf
    Msg = Msg + "Diverse kosten : vanaf " + Left(GrensDetail(3), 7) + " tot " + Right(GrensDetail(3), 7)
    MsgBox Msg
    RekeningOK = False
    Exit Function
Else
    RekeningOK = True
End If
    
End Function

Private Sub Schoon()

    Dim T As Integer
    
    Kontrole.Enabled = False
    uitwisselingOMS = ""
    uitwisselingDATA = ""
    documentLinesDATA = ""
    
    Me.cbImportUBL.Enabled = True

    GeScanBestand = ""
    If StartBlad = True Then
        SSTab2.Tab = 0
        Me.obTAB(0).Value = True
    Else
        SSTab2.Tab = 1
        Me.obTAB(1).Value = True
    End If

    'Me.MiDocView1.filename = ""

    lvLeverancier = ""
    Kontrole.Default = False
    Kontrole.Visible = True
    AankoopDetail.Enabled = False
    'InvoerLijn.Enabled = False
    LeverancierInfo.Caption = ""
    Afsluiten.Visible = False
    Schoonvegen.Enabled = False
    'cmdSwitch.Enabled = False
    'ReturnRooster.Enabled = False
    TekstInfo4.AutoTab = False
    TekstInfo4.ClipMode = 0
    TekstInfo4.PromptInclude = True
    TekstInfo4.Mask = ""

    Err = 0

    TekstInfo0.Enabled = False
    TekstInfo0.Mask = ""
    TekstInfo0.Mask = "##/##/####"
    TekstInfo0.text = MIM_GLOBAL_DATE
    
    TekstInfo1.Enabled = False
    TekstInfo1.Mask = ""
    TekstInfo1.Mask = "##/##/####"
    TekstInfo1.text = MIM_GLOBAL_DATE
    
    TekstInfo2.Enabled = False
    TekstInfo2.Mask = ""
    TekstInfo2.Mask = "##/##/####"
    TekstInfo2.text = MIM_GLOBAL_DATE
    
    TekstInfo3.Enabled = False
    TekstInfo3.text = ""
    TekstInfo3 = LeverancierRekening
    TekstInfo4.Enabled = False
    TekstInfo4.text = ""
    TekstInfo12.Enabled = False
    TekstInfo12.text = ""
    TekstInfo5.Enabled = False
    TekstInfo5.text = ""
    TekstInfo6.Enabled = False
    TekstInfo6.text = ""
    TekstInfo7.Enabled = False
    TekstInfo7.text = ""
    TekstInfo9.Enabled = False
    TekstInfo9.text = ""
    TekstInfo10.Enabled = False
    TekstInfo10.text = ""
    TekstInfo10 = rbtwVAK(4)
    
    RasterSchoon
    If DirekteAankoop.Visible Then
        SSTab1.Tab = 1
    End If
    If Me.obTAB(0).Value = True Then
        SSTab2.Tab = 0
    Else
        SSTab2.Tab = 1
    End If
    Me.cmdXLog.Enabled = False
    Me.LabelInfoXlog.Visible = True

    Me.AankoopOptie(0).Value = True
    
    supplierCompanyId = ""
    supplierCountryCode = ""
    supplierVatNumber = ""
    
End Sub

Private Sub obTAB_Click(Index As Integer)

    BeWaarTekst "DirekteAankoop", "startBlad", Str(obTAB(0).Value)
    
End Sub


Private Sub SchoonVegen_Click()

RasterSchoon
Schoon
SSTab1.SetFocus

End Sub

Function SQLZoekLeverancier()

    If LeverancierInfo.Caption <> "" Then
        KtrlBox = MsgBox("Andere leverancier aanduiden.  Bent U zeker", vbQuestion + vbYesNo + vbDefaultButton2)
        If KtrlBox = vbNo Then
            AankoopDetail.SetFocus
            Exit Function
        End If
    End If
    aIndex = 1
    SharedFl = TABLE_SUPPLIERS
    GridText = ""
    SqlSearch.Show 1
    If Ktrl = 0 Then
        InstalLeverancier
    Else
        Schoon
    End If

End Function

Private Sub SSTab1_Click(PreviousTab As Integer)

    If SSTab1.Tab = 0 And Len(Me.LeverancierInfo) = 0 Then
        SQLZoekLeverancier
    End If

End Sub

Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)

If SSTab1.Tab = 1 Then
    If ListView1.ListItems.Count Then
    ElseIf KeyCode = 13 Then
        KtrlBox = MsgBox("Historiek van reeds bestaande leveranciers tonen", vbQuestion + vbYesNo + vbDefaultButton2)
        If KtrlBox = vbYes Then
            InstalleerRecenteCrediteuren
        End If
    End If
ElseIf KeyCode = 18 Or KeyCode = 13 Then
    SQLZoekLeverancier
End If

End Sub



Private Sub StockBeheer_Click()

On Error Resume Next
AankoopDetail.SetFocus

End Sub



Private Sub TekstInfo0_KeyDown(KeyCode As Integer, Shift As Integer)

    On Local Error Resume Next
    If KeyCode = 13 Then
        SendKeys "{TAB}", True
        Exit Sub
    Else
        TekstInfo0.text = TekstInfo1.text
    End If
    
End Sub

Private Sub TekstInfo0_KeyPress(KeyAscii As Integer)

    TekstInfo0.text = TekstInfo1.text

End Sub


Private Sub TekstInfo1_KeyDown(KeyCode As Integer, Shift As Integer)

    On Local Error Resume Next
    If KeyCode = 13 Then
        SendKeys "{TAB}", True
        Exit Sub
    End If

End Sub

Private Sub TekstInfo10_GotFocus()

    TekstInfo10.SelStart = 0
    TekstInfo10.SelLength = Len(TekstInfo10.text)
    
End Sub



Private Sub TekstInfo12_KeyDown(KeyCode As Integer, Shift As Integer)

    On Local Error Resume Next
    If KeyCode = 13 Then
        SendKeys "{TAB}", True
        Exit Sub
    End If
    
End Sub


Private Sub TekstInfo2_KeyDown(KeyCode As Integer, Shift As Integer)

    On Local Error Resume Next
    If KeyCode = 13 Then
        SendKeys "{TAB}", True
        Exit Sub
    End If
    
End Sub

Private Sub TekstInfo3_GotFocus()

    TekstInfo3.SelStart = 0
    TekstInfo3.SelLength = Len(TekstInfo3.text)
    
End Sub

Private Sub TekstInfo4_KeyDown(KeyCode As Integer, Shift As Integer)

    On Local Error Resume Next
    If KeyCode = 13 Then
        SendKeys "{TAB}", True
        Exit Sub
    End If
    
End Sub

Private Sub TekstInfo5_Change()
    
    If Medekontraktant.Value Then
        TekstInfo7.text = TekstInfo5.text
    End If
    
End Sub

Private Sub TekstInfo5_KeyDown(KeyCode As Integer, Shift As Integer)

    On Local Error Resume Next
    If KeyCode = 13 Then
        SendKeys "{TAB}", True
        Exit Sub
    End If
    
End Sub


Private Sub TekstInfo0_GotFocus()

    TekstInfo0.SelStart = 0
    TekstInfo0.SelLength = Len(TekstInfo0.text)
    TekstInfo0.SelLength = 5
    
End Sub

Private Sub TekstInfo1_GotFocus()

    TekstInfo1.SelStart = 0
    TekstInfo1.SelLength = Len(TekstInfo1.text)
    TekstInfo1.SelLength = 5

End Sub

Private Sub TekstInfo2_GotFocus()

    TekstInfo2.SelStart = 0
    TekstInfo2.SelLength = Len(TekstInfo2.text)
    TekstInfo2.SelLength = 5
    
End Sub

Private Sub TekstInfo6_GotFocus()

    TekstInfo6.SelStart = 0
    TekstInfo6.SelLength = Len(TekstInfo6.text)
    Kontrole.Default = True

End Sub

Private Sub TekstInfo3_KeyDown(KeyCode As Integer, Shift As Integer)

    On Local Error Resume Next
    If KeyCode = 13 Then
        SendKeys "{TAB}", True
        Exit Sub
    End If

    If KeyCode = 17 Then
        aIndex = 0
        SharedFl = TABLE_LEDGERACCOUNTS
        GridText = Trim$(TekstInfo3)
        SqlSearch.Show 1
        If Ktrl = 0 Then
            TekstInfo3 = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
        End If
    End If
    
End Sub

Private Sub TekstInfo10_KeyDown(KeyCode As Integer, Shift As Integer)

    On Local Error Resume Next
    If KeyCode = 13 Then
        SendKeys "{TAB}", True
        Exit Sub
    End If

    If KeyCode = 17 Then
        aIndex = 0
        SharedFl = TABLE_LEDGERACCOUNTS
        GridText = Trim$(TekstInfo10)
        SqlSearch.Show 1
        If Ktrl = 0 Then
            TekstInfo10 = vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #")
        End If
    End If

End Sub


Private Sub TekstInfo3_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub

Private Sub TekstInfo10_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub


Private Sub TekstInfo5_KeyUp(KeyCode As Integer, Shift As Integer)

    Dim RekenHier As Variant

    Select Case KeyCode
        Case 106, 107, 109, 111
            If Len(TekstInfo5) = 1 Then
                Exit Sub
            End If
            RekenHier = LineCalculating(TekstInfo5.text)
            If RekenHier Then
                TekstInfo5.text = Str(Round(RekenHier, 2))
            End If
    End Select

End Sub

Private Sub TekstInfo0_LostFocus()

    If DATE_INVALID((TekstInfo0.text)) Then
        TekstInfo0.text = MIM_GLOBAL_DATE
        TekstInfo0.SetFocus
    ElseIf Not DATE_CHECK((TekstInfo0.text), PERIODAS_TEXT) Then
        BYPERDAT.WindowState = 0
        BYPERDAT.PeriodeBoekjaar.SetFocus
    End If

End Sub

Private Sub TekstInfo1_LostFocus()
    
    If DATE_INVALID((TekstInfo1.text)) Then
        TekstInfo1.text = MIM_GLOBAL_DATE
        TekstInfo1.SetFocus
    Else
        TekstInfo2.text = VValdag((TekstInfo1.text), vBibTekst(TABLE_SUPPLIERS, "#vs04 #"))
        TekstInfo0.text = MIM_GLOBAL_DATE
    End If
    If DATE_KEY(TekstInfo1.text) < Left(BOOKYEAR_FROMTO, 8) Or DATE_KEY(TekstInfo1.text) > Right(BOOKYEAR_FROMTO, 8) Then
        Msg = "Datum aankoopdocument valt BUITEN het actieve boekjaar." & vbCr
        Msg = Msg & "De optie 'boekhoudcontrole' in balans leveranciers zal" & vbCr
        Msg = Msg & "mogelijk niet goed functioneren." & vbCr & vbCr
        Msg = Msg & "Controleer eventueel."
        MsgBox Msg, vbExclamation
    End If
    TekstInfo0.text = TekstInfo1.text

End Sub

Private Sub TekstInfo2_LostFocus()

    If DATE_INVALID((TekstInfo2.text)) Then
        TekstInfo2.text = VValdag((TekstInfo1.text), vBibTekst(TABLE_SUPPLIERS, "#vs04 #"))
        TekstInfo2.SetFocus
    End If

End Sub

Private Sub TekstInfo3_LostFocus()

    If Left(TekstInfo3.text, 3) <> "440" Then
        TekstInfo3.text = LeverancierRekening
        
        TekstInfo3.SetFocus
        Exit Sub
    End If
    bGet TABLE_LEDGERACCOUNTS, 0, vSet((TekstInfo3.text), 7)
    If Ktrl Then
        TekstInfo3.text = LeverancierRekening
        TekstInfo3.SetFocus
    End If
    
End Sub

Private Sub TekstInfo4_LostFocus()
    
    If vBibTekst(TABLE_SUPPLIERS, "#v017 #") = "1" Then
        If BankOk((TekstInfo4.text)) Then
        Else
            On Error Resume Next
            SnelHelpPrint "Gestructureerde betaalreferte onjuist !", BL_LOGGING
            TekstInfo4.text = "OGM onjuist"
            On Error GoTo 0
            Exit Sub
        End If
    Else
        If Trim(TekstInfo4.text) = "" Then
            TekstInfo4.text = TekstInfo12.text
        End If
    End If


End Sub

Private Sub TekstInfo5_LostFocus()

    If AankoopFlg = 1 And AankoopOptie(1).Value = True Then
        If Val(TekstInfo5.text) <> 0 Then
            KtrlBox = MsgBox("De aanbeveling door meeste BTW-diensten om bij creditnota E.U. g��n B.T.W op te nemen negeren", vbQuestion + vbYesNo + vbDefaultButton1)
            If KtrlBox = vbYes Then
                TekstInfo5.text = "0"
            End If
        End If
    ElseIf AankoopFlg = 1 And AankoopOptie(1).Value = False Then
        If Val(TekstInfo5.text) = 0 Then
            MsgBox "Respecteer de aanbeveling door de BTW-diensten om bij factuur E.U. de B.T.W. zelf uit te rekenen en zowel AFTREKBAAR als VERSCHULDIGD (0-operatie dus...) het toepasbaar B.T.W. bedrag mee te delen !", vbExclamation
            TekstInfo5.SetFocus
            Exit Sub
        Else
            TekstInfo7.text = TekstInfo5.text
        End If
    End If

End Sub


Private Sub TekstInfo10_LostFocus()
    
    If Left(TekstInfo10, 6) <> Left(rbtwVAK(4), 6) Then
        TekstInfo10.text = rbtwVAK(4)
        
        TekstInfo10.SetFocus
        Exit Sub
    End If
    bGet TABLE_LEDGERACCOUNTS, 0, vSet((TekstInfo10.text), 7)
    If Ktrl Then
        TekstInfo10.text = rbtwVAK(4)
        TekstInfo10.SetFocus
    End If

End Sub



Private Function WegBoekFout() As Integer

Dim TotaalBedrag    As Currency
Dim Bedrag          As Currency
Dim dInvest         As Currency
Dim dPrive          As Currency
Dim dAlKost         As Currency
Dim dHandel         As Currency

Dim AantalStuks     As Single
Dim T               As Integer

Dim DVT99           As String
Dim LokRekening     As String * 7

bGet TABLE_SUPPLIERS, 0, Left(LeverancierInfo.Caption, 12)
If Ktrl Then
    WegBoekFout = True
    Exit Function
Else
    RecordToVeld TABLE_SUPPLIERS
    dMuntL = Val(TekstInfo9.text)
End If

'gewijzigd voor scanning
WegBoekFout = False
DKTRL_CUMUL = 0: DKTRL_BEF = 0: DKTRL_EUR = 0
Unload frmBoeking
frmBoeking.Hide
TLB_RECORD(TABLE_JOURNAL) = ""
TLB_RECORD(TABLE_INVOICES) = ""
'bijgevoegd voor scanning
If rsMAR(TABLE_INVOICES).State = adStateClosed Then
    Ktrl = bOpen(TABLE_INVOICES)
End If
rsMAR(TABLE_INVOICES).AddNew

vBib TABLE_JOURNAL, "L" + vBibTekst(TABLE_SUPPLIERS, "#A110 #"), "v034"
'19
vBib TABLE_JOURNAL, DATE_KEY((TekstInfo0.text)), "v066"
vBib TABLE_JOURNAL, Document_Key, "v033"
vBib TABLE_JOURNAL, DATE_KEY((TekstInfo1.text)), "v035"
vBib TABLE_JOURNAL, (TekstInfo3.text), "v069"

vBib TABLE_INVOICES, Document_Key, "v033"
vBib TABLE_INVOICES, Document_OGM_NoFormat, "v413"
vBib TABLE_INVOICES, (TekstInfo12.text), "v409"
vBib TABLE_INVOICES, "L" + vBibTekst(TABLE_SUPPLIERS, "#A110 #"), "v034"
vBib TABLE_INVOICES, DATE_KEY((TekstInfo1.text)), "v035"
vBib TABLE_INVOICES, DATE_KEY((TekstInfo2.text)), "v036"
vBib TABLE_INVOICES, (TekstInfo4.text), "v039"
vBib TABLE_INVOICES, Dec$(dMuntL, "###.##########"), "v040"
vBib TABLE_INVOICES, vBibTekst(TABLE_SUPPLIERS, "#vs03 #"), "vs03"
vBib TABLE_JOURNAL, vBibTekst(TABLE_SUPPLIERS, "#A100 #"), "v067"
For T = 0 To AankoopDetail.ListCount - 1
    AankoopDetail.ListIndex = T
    LokRekening = Left(AankoopDetail.text, 7)
    vBib TABLE_JOURNAL, LokRekening, "v019"
    bGet TABLE_LEDGERACCOUNTS, 0, LokRekening
    If Ktrl Then
        WegBoekFout = True
        Exit Function
    End If
    Bedrag = Val(Mid(AankoopDetail.text, 50, 12))
    TotaalBedrag = TotaalBedrag + Bedrag
    If AankoopOptie(0).Value = True Then
        'factuur
        vBib TABLE_JOURNAL, Dec$((Bedrag), MASK_EURBH), "v068"
    ElseIf AankoopOptie(1).Value = True Then
        'creditnota
        vBib TABLE_JOURNAL, Dec$((-Bedrag), MASK_EURBH), "v068"
    Else
        MsgBox "logicafout"
    End If
    
    vBib TABLE_JOURNAL, "", "v102"
    If Len(AankoopDetail.text) > 62 Then
        
        bGet TABLE_PRODUCTS, 0, Mid(AankoopDetail.text, 63, 13)
        If Ktrl Then
            MsgBox "Onverwacht situatie"
            WegBoekFout = True
            Exit Function
        Else
            RecordToVeld TABLE_PRODUCTS
            vBib TABLE_JOURNAL, Mid(AankoopDetail.text, 77, 30), "v067"
            vBib TABLE_JOURNAL, Mid(AankoopDetail.text, 63, 13), "v102"
            AantalStuks = Val(Mid(AankoopDetail.text, 77 + 41, 10))
            If AankoopOptie(0).Value = True Then
                'IseenFactuur
                'AankoopPrijs
                vBib TABLE_PRODUCTS, Str$(Bedrag / AantalStuks), "e113"
                'Leveranciernummer
                vBib TABLE_PRODUCTS, FVT(TABLE_SUPPLIERS, 0), "v124"
                'Aantal stuks aankoop bijvoegen
                vBib TABLE_PRODUCTS, Str$(AantalStuks + Val(vBibTekst(TABLE_PRODUCTS, "#v119 #"))), "v119"
                'Totaal bedrag bijtellen
                vBib TABLE_PRODUCTS, Str$(Bedrag + Val(vBibTekst(TABLE_PRODUCTS, "#e121 #"))), "e121"
                'Laatste kostenrekening onthouden
                vBib TABLE_PRODUCTS, LokRekening, "v116"
            ElseIf AankoopOptie(1).Value = True Then
                'IsCreditnota
                'Aantal stuks aankoop aftrekken
                vBib TABLE_PRODUCTS, Str$(-AantalStuks + Val(vBibTekst(TABLE_PRODUCTS, "#v119 #"))), "v119"
                'Totaal bedrag aftrekken
                vBib TABLE_PRODUCTS, Str$(-Bedrag + Val(vBibTekst(TABLE_PRODUCTS, "#e121 #"))), "e121"
            Else
                MsgBox "logicafout"
            End If
            bUpdate TABLE_PRODUCTS, 0
            If Ktrl Then
                MsgBox "Onverwacht situatie"
                WegBoekFout = True
                Exit Function
            End If
        End If
    End If
    
    'grens 0
    If LokRekening >= Left(GrensDetail(0), 7) And LokRekening <= Right(GrensDetail(0), 7) Then
        dInvest = dInvest + Bedrag
    
    'grens 1
    ElseIf LokRekening >= Left(GrensDetail(1), 7) And LokRekening <= Right(GrensDetail(1), 7) Then
        dPrive = dPrive + Bedrag

    'grens 2
    ElseIf LokRekening >= Left(GrensDetail(2), 7) And LokRekening <= Right(GrensDetail(2), 7) Then
        dHandel = dHandel + Bedrag

    'grens 3
    ElseIf LokRekening >= Left(GrensDetail(3), 7) And LokRekening <= Right(GrensDetail(3), 7) Then
        dAlKost = dAlKost + Bedrag
    Else
        MsgBox "Stop in begrenzing"
    End If
    '1310 IF Left(svkf(1),2)="ON"THEN 1525
    bInsert TABLE_JOURNAL, 0
    If Ktrl Then
        WegBoekFout = True
        Exit Function
    End If
Next
vBib TABLE_JOURNAL, vBibTekst(TABLE_SUPPLIERS, "#A100 #"), "v067"

vBib TABLE_INVOICES, (TekstInfo5.text), "v045"
'If Ar = 3 Then
'    vBib TABLE_INVOICES, "0", "v048"
'    vBib TABLE_INVOICES, "0", "v049"
'    vBib TABLE_INVOICES, "0", "v046"
'    vBib TABLE_INVOICES, "0", "v047"
'Else
    vBib TABLE_INVOICES, Str$(dInvest), "v048"
    vBib TABLE_INVOICES, Str$(dHandel), "v046"
    vBib TABLE_INVOICES, Str$(dAlKost), "v047"
'End If
DVT99 = Str$(dInvest + dAlKost + dHandel + dPrive)

If Medekontraktant.Value Then
    vBib TABLE_INVOICES, (TekstInfo7.text), "v043"
    If Ar = 1 Then
        vBib TABLE_INVOICES, DVT99, "v053"
    Else
        vBib TABLE_INVOICES, DVT99, "v051"
    End If
Else
    If Medekontraktant.Value = 0 And AankoopFlg = 0 Then
        If Ar = 3 Then
            vBib TABLE_INVOICES, DVT99, "v051"
        End If
    End If
End If
If AankoopFlg = 1 Then
    vBib TABLE_INVOICES, (TekstInfo7.text), "v042"
    If Ar = 1 Then
        vBib TABLE_INVOICES, DVT99, "v052"
    Else
        vBib TABLE_INVOICES, DVT99, "v050"
    End If
End If
If AankoopFlg = 2 Then
    vBib TABLE_INVOICES, (TekstInfo7.text), "v044"
    If Ar = 1 Then
        MsgBox "in pseudo vak 88 wordt " & DVT99 & " bijgevoegd.  Bij de eigenlijke BTW aangifte dient U de cijfers van kolom '88!'  bij te tellen bij vak 87", vbInformation
        vBib TABLE_INVOICES, DVT99, "v054"
    Else
        vBib TABLE_INVOICES, DVT99, "v051"
    End If
End If

If Val(TekstInfo5.text) <> 0 Then
    If Ar = 3 Then
        vBib TABLE_JOURNAL, Dec$(-Val(TekstInfo5.text), MASK_EURBH), "v068"
        vBib TABLE_JOURNAL, rbtwVAK(5), "v019"
    Else
        vBib TABLE_JOURNAL, Dec$(Val(TekstInfo5.text), MASK_EURBH), "v068"
        vBib TABLE_JOURNAL, vSet((TekstInfo10.text), 7), "v019"
    End If
    bInsert TABLE_JOURNAL, 0
End If

If AankoopFlg = 0 Then
    If Medekontraktant.Value Then
        If Ar = 1 Then
            vBib TABLE_JOURNAL, Dec$(-Val(TekstInfo7.text), MASK_EURBH), "v068"
        Else
            vBib TABLE_JOURNAL, Dec$(Val(TekstInfo7.text), MASK_EURBH), "v068"
        End If
        FVT(TABLE_JOURNAL, 1) = rbtwVAK(2)
        vBib TABLE_JOURNAL, rbtwVAK(2), "v019"
        bInsert TABLE_JOURNAL, 0
    End If
    GoTo LabeL1440
End If

If AankoopFlg = 2 Then
    If Ar = 1 Then
        vBib TABLE_JOURNAL, Dec$(-Val(TekstInfo7.text), MASK_EURBH), "v068"
    Else
        vBib TABLE_JOURNAL, Dec$(Val(TekstInfo7.text), MASK_EURBH), "v068"
    End If
    vBib TABLE_JOURNAL, rbtwVAK(3), "v019"
Else
    If Ar = 1 Then
        vBib TABLE_JOURNAL, Dec$(-Val(TekstInfo7.text), MASK_EURBH), "v068"
    Else
        vBib TABLE_JOURNAL, Dec$(Val(TekstInfo7.text), MASK_EURBH), "v068"
    End If
    vBib TABLE_JOURNAL, rbtwVAK(1), "v019"
End If
bInsert TABLE_JOURNAL, 0

LabeL1440:
DVT99 = Str$(Val(DVT99) + Val(TekstInfo5.text) - Val(TekstInfo7.text))
If Ar = 1 Then
    vBib TABLE_JOURNAL, Dec$(-Val(DVT99), MASK_EURBH), "v068"
Else
    vBib TABLE_JOURNAL, Dec$(Val(DVT99), MASK_EURBH), "v068"
End If
vBib TABLE_JOURNAL, vSet((TekstInfo3.text), 7), "v019"
bInsert TABLE_JOURNAL, 0

If XisEuroWisBEF = False Then
    vBib TABLE_INVOICES, DVT99, "v249"
Else
    vBib TABLE_INVOICES, Str$(Round(DVT99 / EURO, 2)), "v249"
End If

If GeScanBestand <> "" Then
    FileToBlob rsMAR(TABLE_INVOICES)("bstBLOB37"), GeScanBestand
    rsMAR(TABLE_INVOICES)("bstndNaam37") = GeScanBestand
    rsMAR(TABLE_INVOICES)("typeZending37") = Mid(GeScanBestand, InStr(GeScanBestand, ".") + 1)
End If
bInsert TABLE_INVOICES, 0

If DKTRL_CUMUL <> 0 Then
    frmBoeking.cmdBoeken.Enabled = False
    MsgBox "LogikaFout bij vierkantskontrole journaal." + vbCrLf + vbCrLf + "Deze verrichting wordt geannuleerd.  Controleer zelf eerst en/of raadpleeg ons."
    frmBoeking.Show 1
    WegBoekFout = True
ElseIf JournaalLocked = True Then
    frmBoeking.cmdBoeken.Enabled = False
    frmBoeking.Show 1
    WegBoekFout = True
Else
    Select Case Left(Mim.cmdWegBoekModus.text, 1)
        Case "0"
        Case "1"
            If DKTRL_BEF Or DKTRL_EUR Then
                frmBoeking.Show 1
            End If
        Case "2"
            frmBoeking.Show 1
        Case Else
            MsgBox "situatie..."
    End Select
    If DKTRL_CUMUL Then WegBoekFout = True
End If

End Function
Function InstalleerRecenteCrediteuren()

        Screen.MousePointer = vbHourglass
        Dim rsRecent As ADODB.Recordset
        Set rsRecent = New ADODB.Recordset
        rsRecent.CursorLocation = adUseClient
        'SELECT DISTINCT Leveranciers.A100 AS [Naam], Leveranciers.A110 AS [idCode] FROM Leveranciers, Dokumenten WHERE  'L'+Leveranciers.A110 =  Dokumenten.v034 AND Dokumenten.v033 Like 'A%' ORDER BY Leveranciers.A100
        rsRecent.Open "SELECT DISTINCT Leveranciers.A100 AS [Naam], Leveranciers.A110 AS [idCode] FROM Leveranciers, Dokumenten WHERE  'L'+Leveranciers.A110 =  Dokumenten.v034 AND Dokumenten.v033 Like 'A%' ORDER BY Leveranciers.A100", adntDB, adOpenForwardOnly, adLockReadOnly
        
        Dim itmX As ListItem
        ListView1.ListItems.Clear

        Do While Not rsRecent.EOF
            ' Add a ListItem object.
            Set itmX = ListView1.ListItems.Add(, , rsRecent("Naam"))
                itmX.SubItems(1) = rsRecent("idCode")
            rsRecent.MoveNext
        Loop
        rsRecent.Close
        Set rsRecent = Nothing
        Screen.MousePointer = vbNormal
        ListView1.SetFocus
        
End Function


Private Sub TekstInfo7_KeyDown(KeyCode As Integer, Shift As Integer)

    On Local Error Resume Next
    If KeyCode = 13 Then
        SendKeys "{TAB}", True
        Exit Sub
    End If
    
End Sub





Private Sub TekstInfo9_KeyDown(KeyCode As Integer, Shift As Integer)

    On Local Error Resume Next
    If KeyCode = 13 Then
        SendKeys "{TAB}", True
        Exit Sub
    End If
    
End Sub


Private Sub RefreshReference()
    
    Dim groupCode As String
    
    Select Case Left(Document_Key, 2)
        Case "A0": groupCode = "4"
        Case "A1": groupCode = "3"
    End Select
    
    Dim ReferteTxt As String
    Dim ReferteTxtNoFormat As String
    
    Dim dPip As Double
    Dim recordAsString As String
    Dim chk As String
    
    ReferteTxtNoFormat = groupCode & Mid$(Document_Key, 3, 2) & Mid$(Document_Key, 5, 4) & Mid$(Document_Key, 9) & "xx"
    ReferteTxt = "+++" + groupCode + Mid(Document_Key, 3, 2) + "/" + Mid(Document_Key, 5, 4) + "/" + Mid(Document_Key, 9) + "xx+++"
        
    dPip = Val(Mid$(ReferteTxt, 4, 3) & Mid$(ReferteTxt, 8, 4) & Mid$(ReferteTxt, 13, 3))
    chk = Format$(dPip - Int(dPip / 97) * 97, "00")
    If chk = "00" Then chk = "97"

    Mid$(ReferteTxt, 16, 2) = chk
    Mid$(ReferteTxtNoFormat, 11, 2) = chk
    
    Document_OGM = ReferteTxt
    Document_OGM_NoFormat = ReferteTxtNoFormat
    Me.LabelDocumentReference = "Referte: " & Document_Key & " of " + Document_OGM
    
End Sub



