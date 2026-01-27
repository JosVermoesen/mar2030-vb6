VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form DirekteVerkoop 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7470
   ClientLeft      =   945
   ClientTop       =   1530
   ClientWidth     =   9360
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
   Icon            =   "0400001.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7470
   ScaleWidth      =   9360
   Begin VB.TextBox TextBoxWarningTestCompany 
      BackColor       =   &H000000FF&
      Enabled         =   0   'False
      Height          =   285
      Left            =   6240
      TabIndex        =   64
      TabStop         =   0   'False
      Text            =   "Opgelet: Dit is een testbedrijf!"
      Top             =   6600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame FramePeppol 
      Height          =   735
      Left            =   240
      TabIndex        =   61
      Top             =   5760
      Width           =   7815
      Begin VB.ComboBox cbUncl1001 
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
         Height          =   315
         Left            =   480
         Locked          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   68
         Top             =   240
         Width           =   4095
      End
      Begin VB.CommandButton cbUncl1001Toggle 
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
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   240
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.OptionButton OptionUBL_BE_3_0 
         Caption         =   "UBL.BE 3.0"
         Height          =   375
         Left            =   6360
         TabIndex        =   63
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton OptionPEPPOL_V3 
         Caption         =   "Peppol UBL V3"
         Height          =   375
         Left            =   4680
         TabIndex        =   62
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.CheckBox cbQRCode 
      Caption         =   "QR Betaalcode"
      Height          =   315
      Left            =   6360
      TabIndex        =   59
      Top             =   6960
      Width           =   1695
   End
   Begin VB.CheckBox cbAV 
      Caption         =   "AV toevoegen"
      Height          =   315
      Left            =   4560
      TabIndex        =   58
      Top             =   6960
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CommandButton Annuleren 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   285
      Left            =   8160
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   6960
      Width           =   885
   End
   Begin VB.CommandButton Schoonvegen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Schoo&n"
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
      Height          =   405
      Left            =   8160
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1005
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   120
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   0
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   12091
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
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
      TabPicture(0)   =   "0400001.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(9)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(8)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(6)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblBTWBedrag1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblBTWBedrag2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblBTWBedrag3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(20)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LblInBtw"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LblExBtw"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "LblIn2Btw"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LblEx2Btw"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(5)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(0)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "KlantInfo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "DagKoers"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdSwitch"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TekstInfo3"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "TekstInfo0"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TekstInfo2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "TekstInfo1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "TekstInfo5"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "ckEURINFO"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "chkBTWBouw"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "chkZonderRelatie"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmdLijst"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdSQLInfo"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "CmbExtraAfdruk"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "CmdTekst"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "CmdOmschrijving"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "CmdStock"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Optimaliseer"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Overschrijvingsstrook"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Klassement"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Afsluiten"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Medekontraktant"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "CreditNota"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "KlantAktiveren"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "VerkoopDetail"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Sjabloon"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "VerkoopOptie(0)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "VerkoopOptie(1)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "VerkoopOptie(2)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "sepaDOM"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "cbCheckTools"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "cbMonitortUBL"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "ButtonInfoSupported"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "CheckBoxAlwaysPeppolRefresh"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).ControlCount=   53
      TabCaption(1)   =   "Kettingfacturatie"
      TabPicture(1)   =   "0400001.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvDetail"
      Tab(1).Control(1)=   "cbSelect"
      Tab(1).Control(2)=   "cbFactureren"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Im- en Export"
      TabPicture(2)   =   "0400001.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cbMAPIMAIL"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Plakbord"
      TabPicture(3)   =   "0400001.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lstKopiePlak"
      Tab(3).ControlCount=   1
      Begin VB.CheckBox CheckBoxAlwaysPeppolRefresh 
         Alignment       =   1  'Right Justify
         Caption         =   "Check Peppol Docs"
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
         Left            =   7200
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton ButtonInfoSupported 
         Caption         =   "Info UBL Klant"
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
         Height          =   555
         Left            =   8040
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   5880
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cbMonitortUBL 
         Caption         =   "&UBL B2B OUT"
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
         Height          =   315
         Left            =   5760
         TabIndex        =   65
         ToolTipText     =   "Monitor B2B Verkoopdocumenten"
         Top             =   1755
         Width           =   1455
      End
      Begin VB.CommandButton cbCheckTools 
         Caption         =   "Peppol Tools"
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
         Left            =   5760
         TabIndex        =   60
         Top             =   120
         Width           =   1425
      End
      Begin VB.CheckBox sepaDOM 
         Alignment       =   1  'Right Justify
         Caption         =   "SEPA Domic."
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
         Left            =   8040
         TabIndex        =   57
         Top             =   4440
         Width           =   975
      End
      Begin VB.CommandButton cbMAPIMAIL 
         Caption         =   "mapiMail"
         Height          =   495
         Left            =   -74760
         TabIndex        =   55
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cbFactureren 
         Caption         =   "Factureren"
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
         Left            =   -73560
         TabIndex        =   52
         Top             =   4860
         Width           =   1275
      End
      Begin VB.CommandButton cbSelect 
         Caption         =   "(De)Selecteren"
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
         Left            =   -74880
         TabIndex        =   51
         Top             =   4860
         Width           =   1275
      End
      Begin VB.ListBox lstKopiePlak 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5310
         Left            =   -74940
         TabIndex        =   49
         Top             =   60
         Width           =   8715
      End
      Begin VB.OptionButton VerkoopOptie 
         Alignment       =   1  'Right Justify
         Caption         =   "Offerte"
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
         Left            =   8010
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   660
         Width           =   1005
      End
      Begin VB.OptionButton VerkoopOptie 
         Alignment       =   1  'Right Justify
         Caption         =   "Bestelbon"
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
         Left            =   7770
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   420
         Width           =   1245
      End
      Begin VB.OptionButton VerkoopOptie 
         Alignment       =   1  'Right Justify
         Caption         =   "&Directe Verkoop"
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
         Left            =   7290
         TabIndex        =   19
         Top             =   180
         Value           =   -1  'True
         Width           =   1725
      End
      Begin VB.CommandButton Sjabloon 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sja&bloon"
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
         Left            =   5760
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   480
         Width           =   1425
      End
      Begin VB.ListBox VerkoopDetail 
         BackColor       =   &H00FFFFFF&
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
         Height          =   2580
         Left            =   0
         TabIndex        =   17
         Top             =   2400
         Width           =   7995
      End
      Begin VB.CommandButton KlantAktiveren 
         Caption         =   "&Klant"
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
         Left            =   3840
         TabIndex        =   0
         Top             =   1320
         Width           =   1125
      End
      Begin VB.CheckBox CreditNota 
         Alignment       =   1  'Right Justify
         Caption         =   "&CreditNota"
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
         Left            =   7620
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   900
         Width           =   1395
      End
      Begin VB.CheckBox Medekontraktant 
         Alignment       =   1  'Right Justify
         Caption         =   "Bt&w Medekontractant"
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
         Left            =   7170
         TabIndex        =   15
         Top             =   1560
         Width           =   1845
      End
      Begin VB.CommandButton Afsluiten 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Verwerken"
         Enabled         =   0   'False
         Height          =   330
         Left            =   3840
         TabIndex        =   14
         Top             =   1740
         Width           =   1125
      End
      Begin VB.CommandButton Klassement 
         Caption         =   "&Haal uit klassement"
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
         Height          =   555
         Left            =   3840
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   720
         Width           =   1125
      End
      Begin VB.CheckBox Overschrijvingsstrook 
         Alignment       =   1  'Right Justify
         Caption         =   "&Met OVS Strook"
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
         Left            =   8040
         TabIndex        =   12
         Top             =   3840
         Width           =   975
      End
      Begin VB.CommandButton Optimaliseer 
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
         Height          =   570
         Left            =   8040
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3120
         Width           =   945
      End
      Begin VB.CommandButton CmdStock 
         Caption         =   "&Stock"
         Height          =   330
         Left            =   0
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1740
         Width           =   915
      End
      Begin VB.CommandButton CmdOmschrijving 
         Caption         =   "Via &Omschrijving"
         Height          =   330
         Left            =   960
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1740
         Width           =   1605
      End
      Begin VB.CommandButton CmdTekst 
         Caption         =   "Vrije &Tekst"
         Height          =   330
         Left            =   2640
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1740
         Width           =   1065
      End
      Begin VB.CommandButton CmbExtraAfdruk 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "(Extra) Afdruk"
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
         Height          =   525
         Left            =   8040
         TabIndex        =   7
         Top             =   2400
         Width           =   945
      End
      Begin VB.CommandButton cmdSQLInfo 
         Caption         =   "SQL &Info"
         Height          =   525
         Left            =   5040
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   720
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.CommandButton cmdLijst 
         Caption         =   "Totaal &Lijst"
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
         Left            =   5760
         TabIndex        =   5
         Top             =   840
         Width           =   1425
      End
      Begin VB.CheckBox chkZonderRelatie 
         Caption         =   "&Te factureren"
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
         Left            =   5760
         TabIndex        =   4
         Top             =   1215
         Value           =   1  'Checked
         Width           =   1275
      End
      Begin VB.CheckBox chkBTWBouw 
         Alignment       =   1  'Right Justify
         Caption         =   "BtwBouw + 5 jaar"
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
         Left            =   7410
         TabIndex        =   3
         Top             =   1800
         Width           =   1605
      End
      Begin VB.CheckBox ckEURINFO 
         Alignment       =   1  'Right Justify
         Caption         =   "&Euro && Bef Info"
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
         Left            =   3420
         TabIndex        =   2
         Top             =   2100
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin MSMask.MaskEdBox TekstInfo5 
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   65535
         Enabled         =   0   'False
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
      Begin MSMask.MaskEdBox TekstInfo1 
         Height          =   255
         Left            =   2940
         TabIndex        =   23
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   65535
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo2 
         Height          =   285
         Left            =   4380
         TabIndex        =   24
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   65535
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
      Begin MSMask.MaskEdBox TekstInfo0 
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   65535
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox TekstInfo3 
         Height          =   285
         Left            =   2040
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2100
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   65535
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ListView lvDetail 
         Height          =   4815
         Left            =   -75000
         TabIndex        =   50
         Top             =   0
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   8493
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   12648384
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
      Begin VB.CommandButton cmdSwitch 
         Caption         =   "Ingave in BEF"
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
         Height          =   555
         Left            =   4680
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   4080
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label DagKoers 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Dagkoers"
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
         Left            =   600
         TabIndex        =   47
         Top             =   60
         Width           =   750
      End
      Begin VB.Label KlantInfo 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Height          =   1035
         Left            =   0
         TabIndex        =   46
         Top             =   660
         Width           =   3765
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prijs"
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
         Left            =   4980
         TabIndex        =   45
         Top             =   2100
         Width           =   465
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aantal"
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
         Left            =   5700
         TabIndex        =   44
         Top             =   2100
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LijnTotaal"
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
         Left            =   6720
         TabIndex        =   43
         Top             =   2100
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "40Rekening"
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
         Left            =   4440
         TabIndex        =   42
         Top             =   60
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Vervaldag"
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
         Left            =   2940
         TabIndex        =   41
         Top             =   60
         Width           =   780
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Verkoopdetail lijnen"
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
         Left            =   0
         TabIndex        =   40
         Top             =   2100
         Width           =   1515
      End
      Begin VB.Label LblEx2Btw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6540
         TabIndex        =   39
         ToolTipText     =   "Totaal dokument in EURO exclusief BTW"
         Top             =   5340
         Width           =   1365
      End
      Begin VB.Label LblIn2Btw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   38
         ToolTipText     =   "Totaal dokument in EURO inclusief BTW"
         Top             =   5340
         Width           =   1365
      End
      Begin VB.Label LblExBtw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6540
         TabIndex        =   37
         ToolTipText     =   "Totaal dokument in BEF exclusief BTW"
         Top             =   5040
         Width           =   1365
      End
      Begin VB.Label LblInBtw 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   36
         ToolTipText     =   "Totaal dokument in BEF inclusief BTW"
         Top             =   5040
         Width           =   1365
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EUR incl./excl. Btw"
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
         Left            =   3660
         TabIndex        =   35
         Top             =   5340
         Width           =   1485
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BEF incl./excl. Btw"
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
         Left            =   3660
         TabIndex        =   34
         Top             =   5040
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Datum Document"
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
         Left            =   1440
         TabIndex        =   33
         Top             =   60
         Width           =   1305
      End
      Begin VB.Label lblBTWBedrag3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   32
         ToolTipText     =   "Bij afdruk/boeking nog te manipuleren via vlag medekontraktant indien nodig"
         Top             =   5340
         Width           =   1245
      End
      Begin VB.Label lblBTWBedrag2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1140
         TabIndex        =   31
         ToolTipText     =   "Bij afdruk/boeking nog te manipuleren via vlag medekontraktant indien nodig"
         Top             =   5340
         Width           =   1245
      End
      Begin VB.Label lblBTWBedrag1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -120
         TabIndex        =   30
         ToolTipText     =   "Bij afdruk/boeking nog te manipuleren via vlag medekontraktant indien nodig"
         Top             =   5340
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Btw Basis 1"
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
         Left            =   -120
         TabIndex        =   29
         ToolTipText     =   "Bij afdruk/boeking nog te manipuleren via vlag medekontraktant indien nodig"
         Top             =   5040
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Btw Basis 2"
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
         Left            =   1140
         TabIndex        =   28
         ToolTipText     =   "Bij afdruk/boeking nog te manipuleren via vlag medekontraktant indien nodig"
         Top             =   5040
         Width           =   1245
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Btw Basis 3"
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
         Left            =   2400
         TabIndex        =   27
         ToolTipText     =   "Bij afdruk/boeking nog te manipuleren via vlag medekontraktant indien nodig"
         Top             =   5040
         Width           =   1245
      End
   End
   Begin VB.CommandButton cbLayOudPDF 
      Caption         =   "Beheer van VPE Lay-Out"
      Height          =   315
      Left            =   120
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   6960
      Width           =   2775
   End
   Begin MSMAPI.MAPIMessages MPIBericht 
      Left            =   600
      Top             =   0
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
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "DirekteVerkoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefInt A-Z

Dim reprintOnly As Boolean

Dim rsSellerUBL As New ADODB.Recordset

Dim adminNoVat As Boolean
Dim orderMarReferences As String

Dim idPdfForUbl As String
Dim printAndSave As Boolean

Dim path As String * 260
Dim ret As Integer
Dim desktopLocatie As String

Dim emailTemplate As String
Dim sellerInfo As String
                
'UBL
Dim base64Str As String
Dim documentTemplate As String
Dim invoiceLineTemplate As String
Dim invoiceTaxLineTemplate As String
Dim taxSubtotalLine As String
Dim listTaxSubtotalLines As String
Dim thisInvoiceLine As String
Dim listInvoiceLines As String
Dim tmpString As String

'SUPPLIER:
Dim invoiceNumber As String '240100225
Dim invoiceCustomerNumber As String ' 844
Dim invoiceCurrency As String 'EUR
Dim invoiceDate As String '2024-01-31
Dim invoiceDueDate As String '2024-03-01
Dim invoiceTypeCode As String '380 (see https://docs.peppol.eu/poac/pint/pint/trn-invoice/codelist/UNCL1001-inv/)
Dim orderReference As String 'ex. 024010022505 as of +++024/0100/22505+++
Dim supplierTaxScheme As String 'ex. VAT
Dim supplierRegistrationId As String 'ex. 0423100736
Dim supplierVatNumber As String 'ex. 0423100736 (Titeca)
Dim supplierRegistrationName As String 'Titeca Accountancy Merelbeke NV
Dim supplierStreetName As String 'Fraterstraat 132
Dim supplierCityName As String 'MERELBEKE
Dim supplierPostalZone As String '9820
Dim supplierCountryCode As String 'BE
Dim supplierCompanyId As String '0423100736
Dim supplierCompanyIdExtended As String 'BE0423100736
Dim supplierTelephone As String '09 232 28 00

'CUSTOMER:
Dim customerVatNumber As String '0440058217
Dim customerTaxScheme As String 'ex. VAT
'if customer is VAT then else ?
Dim customerName As String 'ex. Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv
Dim customerStreetName As String 'Grote Baan 141
Dim customerCityName As String 'HERDERSEM
Dim customerPostalZone As String '9310
Dim customerCountryCode As String 'BE
Dim customerCompanyId As String '04400580217
Dim customerCompanyIdExtended As String ''BE04400580217
Dim customerSchemeId As String '0208
Dim customerRegistrationName As String 'Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv
Dim customerElectronicMail As String 'info@rv.be

'PAYMENTMEANS/TERMS
Dim paymentMeansCode As String '1 see: https://docs.peppol.eu/poac/pint/pint/trn-invoice/codelist/UNCL4461/
Dim paymentID As String '024010022505
Dim supplierIBAN As String 'BE50001683378618
Dim supplierBIC As String 'GEBABEBB
Dim paymentTerms As String '024010022505

'TAXTOTAL
Dim subTaxTotalAmount As String '435.86
Dim subTaxableAmount As String '2075.50
Dim subTaxAmount As String '435.86
Dim subTaxCategory As String '03
Dim subTaxPercent As String '21.00
Dim subTaxExclusiveAmount As String
Dim subTaxInclusiveAmount As String
Dim subPayableAmount As String

Dim taxGlobalTotalAmount As String
Dim taxableGlobalTotalAmount As String
Dim globalTotalAmount As String

'LEGALMTOTAL
Dim taxExclusiveAmount As String '2075.50
Dim taxInclusiveAmount As String '2511.36
Dim payableAmount As String '2511.36

'LINES
Dim lineCounter As String
Dim lineQuantity As String
Dim lineExtentionAmount As String
Dim lineAccountingcost As String
Dim lineTaxAmountAccurate As String
Dim lineDescription As String
Dim lineSellerItemIdentification As String
Dim lineTaxCategory As String
Dim lineTaxPercentage As String
Dim linePriceAmount As String

Dim vatGroupsCount As Integer
Dim ublTT As Integer
Dim dVeldUblInfo(7) As Double
Dim VeldUblInfo(9) As String


Dim sy                  As String
Dim sy2                 As String

Dim dokumentType         As String * 2
Dim pdfDOKUMENTTYPE As String
Dim dokumentHistoriek As String * 11
Dim T As Integer

Dim VerkoopDetailTitel(8) As String
Dim dokumentSleutel As String * 11
Dim KlantRekening As String * 7

Dim VerkoopFLG As Integer
Dim AfdrukFlag As Integer
Dim AantalEx As String * 2
Dim DefaultVerkoop As String * 7
Dim KontaktPersoon As Integer
Dim rbtwVAK(10) As String * 7
Dim Vr As Integer
Dim sMuntKlant As String * 3

Dim BTWEuroBasis(3) As Currency
Dim BTWEuroBedrag(3) As Currency

Dim BTWBasis(3) As Currency
Dim BTWBedrag(3) As Currency
Dim TotaalBTW As Currency
Dim TotaalUitvoer As Currency

Dim dMuntK              As Double
Dim sMunt               As String * 3

Dim VsoftVanaf          As String
Dim VsoftTot            As String
Dim BeginXbox           As Integer
Dim BeginYbox           As Integer
Dim OVSStrooklijnen As Long

Dim ForFait             As Integer
Dim SteedsDrukken As Boolean

Dim vkMaskAantal As String

Dim rsDetail    As New ADODB.Recordset
Dim rsKlant     As New ADODB.Recordset

Dim pdfX As Double
Dim pdfY As Double
Dim pdfOVSStrook As Double
Dim isDOM As Boolean

Dim locPOSTVAKIN As String
Dim customerPrefersEmail As Boolean
Dim useEmail As Boolean
Dim customerA110 As String

Function checkForB2BInvoice() As Boolean

    checkForB2BInvoice = False
    Me.FramePeppol.Visible = False
    
    customerElectronicMail = Trim(RV(rsKlant, "v224"))
    If Trim(customerElectronicMail) = "" Then
        MsgBox "Straks een B2B klant met een Peppol document. Adres en/of mailadres ontbreken." & vbCrLf & vbCrLf & "Klantfiche eerst bijwerken a.u.b.!", vbCritical
        Exit Function
    End If
    
    customerCountryCode = Trim(RV(rsKlant, "v150"))
    If customerCountryCode = "BE" Then
        If BtwKontrole(customerVatNumber, True) = customerVatNumber Then
        Else
            MsgBox "Deze klant heeft een ongeldig Belgisch Btwnummer." & vbCrLf & vbCrLf & "Klantfiche eerst bijwerken a.u.b.!", vbCritical
            Exit Function
        End If
        If BtwKontrole(customerCompanyId, True) = customerCompanyId Then
        Else
            MsgBox "Deze klant heeft een ongeldig Belgisch Ondernemingsnummer." & vbCrLf & vbCrLf & "Klantfiche eerst bijwerken a.u.b.!", vbCritical
            Exit Function
        End If
        
        If Trim(RV(rsKlant, "v407")) = "" Then
            Msg = "Geen Peppol ondersteunde documenten gevonden in deze klant zijn fiche" & vbCrLf & vbCrLf
            Msg = Msg & "Is marIntegraal reeds geactiveerd voor Peppol uitwisseling?" & vbCrLf
            Msg = Msg & "Vernieuw dan met MarSync eerst uw klantenfiches a.u.b." & vbCrLf & vbCrLf
            Msg = Msg & "Hierna wordt op klassieke wijze verder gewerkt." & vbCrLf & vbCrLf
            Msg = Msg & "Na 29/12/2025 zijn voor uw Belgische B2B klanten enkel nog verzendingen Peppol factuur of creditnota mogelijk"
            MsgBox Msg, vbInformation
            Exit Function
        Else
            'urn:fdc:peppol.eu:2017:poacc:billing:3.0::2.1
            If InStr(RV(rsKlant, "v407"), "urn:fdc:peppol.eu:2017:poacc:billing:3.0::2.1") Then
                Me.OptionPEPPOL_V3.Enabled = True
                Me.OptionPEPPOL_V3.Value = vbChecked
            ElseIf TextBoxWarningTestCompany.Visible Then
                Me.OptionPEPPOL_V3.Enabled = True
            Else
                Me.OptionPEPPOL_V3.Enabled = False
            End If
            'urn:UBL.BE:1.0.0.20180214::2.1
            If InStr(RV(rsKlant, "v407"), "urn:UBL.BE:1.0.0.20180214::2.1") Then
                Me.OptionUBL_BE_3_0.Enabled = True
                Me.OptionUBL_BE_3_0.Value = True
            ElseIf TextBoxWarningTestCompany.Visible Then
                Me.OptionUBL_BE_3_0.Enabled = True
            Else
                Me.OptionUBL_BE_3_0.Enabled = False
            End If
            If Me.OptionPEPPOL_V3.Enabled = False Then
                Msg = "Geen Peppol ondersteunde documenten gevonden in deze klant zijn fiche" & vbCrLf & vbCrLf
                Msg = Msg & "Vernieuw eerst de klantfiche via knop [Peppol Tools] a.u.b." & vbCrLf & vbCrLf
                Msg = Msg & "Hierna wordt op klassieke wijze verder gewerkt (indien klant niet geregistreerd)." & vbCrLf & vbCrLf
                Msg = Msg & "Vanaf 01/03/2026 zijn voor uw Belgische B2B klanten strikte regels en mag U enkel nog B2B factuur of creditnota verzendingen langsom Peppol"
                MsgBox Msg, vbExclamation
                Exit Function
            ElseIf TextBoxWarningTestCompany.Visible Then
                Msg = "Gelieve in dit testbedrijf enkel de klant Roelandt & Vermoesen te gebruiken" & vbCrLf & vbCrLf
                Msg = Msg & "Gelieve geen extra klanten toe te voegen." & vbCrLf & vbCrLf
                Msg = Msg & "Vernieuw met MarSync indien extra ondersteunde documenten toegevoegd werden."
                MsgBox Msg, vbInformation
            Else
                Me.ButtonInfoSupported.Enabled = True
                Me.ButtonInfoSupported.Visible = True
            End If
        End If
    End If
    
    Select Case customerCountryCode
        Case "AT"
            customerSchemeId = "9914"
        Case "BE"
            customerSchemeId = "0208"
        Case "BG"
            customerSchemeId = "9910"
        Case "HR"
            customerSchemeId = "9916"
        Case "CY"
            customerSchemeId = "9915"
        Case "CZ"
            customerSchemeId = "9913"
        Case "DK"
            customerSchemeId = "9902"
        Case "EE"
            customerSchemeId = "9917"
        Case "FI"
            customerSchemeId = "0037"
        Case "FR"
            customerSchemeId = "0002"
        Case "DE"
            customerSchemeId = "0204"
        Case "GR"
            customerSchemeId = "9918"
        Case "HU"
            customerSchemeId = "9919"
        Case "IS"
            customerSchemeId = "9908"
        Case "IE"
            customerSchemeId = "9920"
        Case "IT"
            customerSchemeId = "0201"
        Case "LV"
            customerSchemeId = "9921"
        Case "LT"
            customerSchemeId = "9922"
        Case "LU"
            customerSchemeId = "9923"
        Case "MT"
            customerSchemeId = "9924"
        Case "NL"
            customerSchemeId = "0106"
        Case "NO"
            customerSchemeId = "9908"
        Case "PL"
            customerSchemeId = "9926"
        Case "PT"
            customerSchemeId = "9927"
        Case "RO"
            customerSchemeId = "9928"
        Case "SK"
            customerSchemeId = "9929"
        Case "SI"
            customerSchemeId = "9930"
        Case "ES"
            customerSchemeId = "9957"
        Case "SE"
            customerSchemeId = "0007"
        Case "CH"
            customerSchemeId = "9955"
        Case "GB"
            customerSchemeId = "9932"
        Case Else
            Msg = "marIntegraal geeft nog geen Peppol ondersteuning voor " & customerCountryCode & vbCrLf & vbCrLf
            Msg = Msg & "Verzendt verkoopdocumenten op de klassieke wijze"
            MsgBox Msg, vbInformation
            Exit Function
    End Select
    
    If Me.VerkoopOptie(0).Value Or Me.CreditNota.Value Then
        'set Commercial invoice default
        Me.cbCheckTools.Enabled = True
        Me.cbUncl1001.Visible = True
        Me.cbUncl1001Toggle.Visible = True
        Me.FramePeppol.Visible = True
        Me.OptionPEPPOL_V3.Visible = True
        Me.OptionUBL_BE_3_0.Visible = True
    
        Me.cbUncl1001.ListIndex = 0
        'Me.OptionUBL_BE_3_0.Value = True
        'Me.OptionPEPPOL_V3.Value = True
        'invoiceTypeCode = "380"
    Else
        Me.cbCheckTools.Enabled = False
        Me.cbUncl1001.Visible = False
        Me.cbUncl1001Toggle.Visible = False
        Me.FramePeppol.Visible = False
        Me.OptionPEPPOL_V3.Visible = False
        Me.OptionUBL_BE_3_0.Visible = False
        Me.OptionUBL_BE_3_0.Value = False
    End If
    'check invoice or creditnote
    If Me.CreditNota.Value Then
        Me.cbUncl1001.ListIndex = 1
    Else
        Me.cbUncl1001.ListIndex = 0
    End If
    checkForB2BInvoice = True
    
End Function

Sub LaadHetdokument()
Dim aa As String
Dim T As Integer

If dokumentType = "15" Then
    RasterSchoon
End If

Set rsDetail = New ADODB.Recordset
    
On Error Resume Next
Err = 0
rsDetail.CursorLocation = adUseClient
Msg = "SELECT * FROM Allerlei WHERE v005 Like '" + dokumentType + Left(XLogKey, 11) + "%'"
SnelHelpPrint Msg, BL_LOGGING
Screen.MousePointer = vbHourglass
rsDetail.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Screen.MousePointer = vbNormal
    Exit Sub
ElseIf rsDetail.RecordCount = 0 Then
    MsgBox "Geen dokumenten (meer) te vinden.", vbInformation
    Screen.MousePointer = vbNormal
    Exit Sub
End If

If InStr(ADOBIB_TEXT(rsDetail("Memo"), "#v147 #"), "V") Then
    Msg = "dokument: " + Left(XLogKey, 11)
    Msg = Msg + " reeds in relatie tot " + vbCrLf
    Msg = Msg + "verkoopdokument(en): " + ADOBIB_TEXT(rsDetail("Memo"), "#v147 #") + vbCrLf + vbCrLf
    Msg = Msg + "Toch ophalen ?"
    KtrlBox = MsgBox(Msg, vbExclamation + vbYesNo + vbDefaultButton2, "Opletten !")
    If KtrlBox = vbNo Then
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
End If

dokumentSleutel = ADOBIB_TEXT(rsDetail("Memo"), "#v033 #")
If Mid(dokumentSleutel, 2, 1) = "1" Then
    CreditNota.Value = 1
End If

Dim docInEur As String

TekstInfo3 = dokumentSleutel
Medekontraktant.Value = Val(ADOBIB_TEXT(rsDetail("Memo"), "#v135 #"))
TekstInfo0.Text = DATE_TEXT(ADOBIB_TEXT(rsDetail("Memo"), "#v035 #"))
TekstInfo1.Text = DATE_TEXT(ADOBIB_TEXT(rsDetail("Memo"), "#v036 #"))
TekstInfo2.Text = ADOBIB_TEXT(rsDetail("Memo"), "#v136 #")
DirekteVerkoop.Caption = Left(DirekteVerkoop.Caption, 28) + "(" + dokumentSleutel + ")"
docInEur = ADOBIB_TEXT(rsDetail("Memo"), "#vEUR #")

Do While Not rsDetail.EOF
    GoSub VoegDetailIn
    rsDetail.MoveNext
Loop
rsDetail.Close
Set rsDetail = Nothing
Annuleren.Enabled = False
Afsluiten.Enabled = True
Me.CmbExtraAfdruk.Enabled = True
On Local Error Resume Next
If docInEur = "EUR" Then
    cmdSwitch.Caption = "Ingave in EUR"
Else
    cmdSwitch.Caption = "Ingave in BEF"
End If

Afsluiten.SetFocus
Select Case dokumentType
    Case "15"
        For T = 0 To 2
            VerkoopOptie(T).Enabled = False
        Next
        TekstInfo0.Enabled = False
        TekstInfo1.Enabled = False
        TekstInfo2.Enabled = False
        
        Medekontraktant.Enabled = False
        CreditNota.Enabled = False
        VerkoopDetail.Enabled = False
        Afsluiten.Enabled = False

    Case "14"
        VerkoopOptie(2).Enabled = False
        frmDokHistoriek.lstDokHistoriek.AddItem dokumentSleutel
        If frmDokHistoriek.Visible = False Then
            frmDokHistoriek.Show
        End If
        DirekteVerkoop.SetFocus
                                    
    Case "13"
        frmDokHistoriek.lstDokHistoriek.AddItem dokumentSleutel
        If frmDokHistoriek.Visible = False Then
            frmDokHistoriek.Show
        End If
        DirekteVerkoop.SetFocus
          
End Select
RefreshBTW
Screen.MousePointer = vbNormal
Exit Sub

VoegDetailIn:
For T = 0 To 7
    aa = ADOBIB_TEXT(rsDetail("Memo"), "#v" + Format(139 + T, "000") + " #")
    If aa <> "" Then
        VerkoopDetail.AddItem aa
    End If
Next
Return

End Sub

Function XmlInvoiceGenerateUBLBE30NoVat(thisPdf As String) As Boolean

'UBL
'base64Str = ""
documentTemplate = ""
invoiceLineTemplate = ""
invoiceTaxLineTemplate = ""
taxSubtotalLine = ""
listTaxSubtotalLines = ""
thisInvoiceLine = ""
listInvoiceLines = ""
tmpString = ""
'SUPPLIER:
'invoiceNumber = "" '240100225
'invoiceCustomerNumber = "" ' 844
invoiceCurrency = "" 'EUR
'invoiceDate = "" '2024-01-31
'invoiceDueDate = "" '2024-03-01
'invoiceTypeCode = "" '380 (see https://docs.peppol.eu/poac/pint/pint/trn-invoice/codelist/UNCL1001-inv/)
'orderReference = "" 'ex. 024010022505 as of +++024/0100/22505+++
'supplierTaxScheme = "" 'ex. VAT
'supplierRegistrationId = "" 'ex. 0423100736
'supplierVatNumber = "" 'ex. 0423100736 (Titeca)
'supplierRegistrationName = "" 'Titeca Accountancy Merelbeke NV
'supplierStreetName = "" 'Fraterstraat 132
'supplierCityName = "" 'MERELBEKE
'supplierPostalZone = "" '9820
'supplierCountryCode = "" 'BE
'supplierCompanyId = "" 'BE0423100736
'supplierTelephone = "" '09 232 28 00
'CUSTOMER:
'customerVatNumber = "" '0440058217
'customerTaxScheme = "" 'ex. VAT
'if customer is VAT then else ?
'customerName = "" 'ex. Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv
'customerStreetName = "" 'Grote Baan 141
'customerCityName = "" 'HERDERSEM
'customerPostalZone = "" '9310
'customerCountryCode = "" 'BE
'customerCompanyId = "" 'BE04400580217
'customerRegistrationName = "" 'Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv
'customerElectronicMail = "" 'info@rv.be
'PAYMENTMEANS/TERMS
paymentMeansCode = "1" '1 see: https://docs.peppol.eu/poac/pint/pint/trn-invoice/codelist/UNCL4461/
'paymentID = "" '024010022505
'supplierIBAN = "" 'BE50001683378618
'supplierBIC = "" 'GEBABEBB
'paymentTerms = "" '024010022505
'TAXTOTAL
subTaxTotalAmount = "" '435.86
subTaxableAmount = "" '2075.50
subTaxAmount = "" '435.86
subTaxCategory = "" '03
subTaxPercent = "" '21.00
subTaxExclusiveAmount = ""
subTaxInclusiveAmount = ""
subPayableAmount = ""

'taxGlobalTotalAmount = ""
'taxableGlobalTotalAmount = ""
globalTotalAmount = ""

'LEGALMTOTAL
taxExclusiveAmount = "" '2075.50
taxInclusiveAmount = "" '2511.36
payableAmount = "" '2511.36

'LINES
lineCounter = ""
lineQuantity = ""
lineExtentionAmount = ""
lineAccountingcost = ""
lineTaxAmountAccurate = ""
lineDescription = ""
lineSellerItemIdentification = ""
lineTaxCategory = ""
lineTaxPercentage = ""
linePriceAmount = ""
vatGroupsCount = 0
ublTT = 0

XmlInvoiceGenerateUBLBE30NoVat = False
If Me.OptionUBL_BE_3_0.Enabled = False Then
    Exit Function
End If

Ktrl = ScrLeesBestandAlleTekst(documentTemplate, PROGRAM_LOCATION + "xml-templates\ubl_be_3_0-invoice-no-vat.xml")
If Ktrl = 0 Then
    MsgBox "Onverwachte situatie", vbCritical
End If

'invoiceNumber = "240100225"
documentTemplate = Replace(documentTemplate, "{documentNumber}", invoiceNumber)
'invoiceCustomerNumber = "844"
documentTemplate = Replace(documentTemplate, "{customerNumber}", invoiceCustomerNumber)
invoiceCurrency = "EUR"
documentTemplate = Replace(documentTemplate, "{currency}", invoiceCurrency)

If orderMarReferences = "" Then
    documentTemplate = Replace(documentTemplate, "{orderReference}", invoiceNumber)
Else
    documentTemplate = Replace(documentTemplate, "{orderReference}", orderMarReferences)
End If

base64Str = GetBase64Str(thisPdf)
documentTemplate = Replace(documentTemplate, "{pdfFileName}", invoiceNumber + ".pdf")
documentTemplate = Replace(documentTemplate, "{pdfBase64}", base64Str)
    
'invoiceDate = "2024-01-31"
documentTemplate = Replace(documentTemplate, "{documentDate}", invoiceDate)
'invoiceDueDate = "2024-03-01"
documentTemplate = Replace(documentTemplate, "{documentDueDate}", invoiceDueDate)
documentTemplate = Replace(documentTemplate, "{documentTypeCode}", invoiceTypeCode)
'orderReference = "024010022505" 'as of +++024/0100/22505+++
documentTemplate = Replace(documentTemplate, "{orderReference}", orderReference)
documentTemplate = Replace(documentTemplate, "{supplierTaxScheme}", supplierTaxScheme)
'supplierVatNumber = "0423100736"
documentTemplate = Replace(documentTemplate, "{supplierVatNumber}", supplierVatNumber)
'tmpString = String99(READING, 46)
'If InStr(tmpString, "&") Then 'verbeteren voor XML bestand!!!
'    tmpString = Replace(tmpString, "&", "&amp;")
'End If
'supplierRegistrationId = "0423100736"
documentTemplate = Replace(documentTemplate, "{supplierRegistrationId}", supplierRegistrationId)
'supplierRegistrationName = "Titeca Accountancy Merelbeke NV"
documentTemplate = Replace(documentTemplate, "{supplierRegistrationName}", supplierRegistrationName)
'supplierStreetName = "Fraterstraat 132"
documentTemplate = Replace(documentTemplate, "{supplierStreetName}", supplierStreetName)
'supplierCityName = "MERELBEKE"
documentTemplate = Replace(documentTemplate, "{supplierCityName}", supplierCityName)
'supplierPostalZone = "9820"
documentTemplate = Replace(documentTemplate, "{supplierPostalZone}", supplierPostalZone)
'supplierCountryCode = "BE"
documentTemplate = Replace(documentTemplate, "{supplierCountryCode}", supplierCountryCode)
'supplierCompanyId = "0423100736"
documentTemplate = Replace(documentTemplate, "{supplierCompanyId}", supplierCompanyId)
'supplierCompanyIdExtended = "BE0423100736"
documentTemplate = Replace(documentTemplate, "{supplierCompanyIdExtended}", supplierCompanyIdExtended)

'supplierTelephone = "09 232 28 00"
documentTemplate = Replace(documentTemplate, "{supplierTelephone}", supplierTelephone)

'customerVatNumber = "0440058217"
documentTemplate = Replace(documentTemplate, "{customerVatNumber}", customerVatNumber)
documentTemplate = Replace(documentTemplate, "{customerTaxScheme}", customerTaxScheme)
'customerName = "Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv"
documentTemplate = Replace(documentTemplate, "{customerRegistrationName}", customerName)
'customerStreetName = "Grote Baan 141"
documentTemplate = Replace(documentTemplate, "{customerStreetName}", customerStreetName)
'customerCityName = "HERDERSEM"
documentTemplate = Replace(documentTemplate, "{customerCityName}", customerCityName)
'customerPostalZone = "9310"
documentTemplate = Replace(documentTemplate, "{customerPostalZone}", customerPostalZone)
'customerCountryCode = "BE"
documentTemplate = Replace(documentTemplate, "{customerCountryCode}", customerCountryCode)
'customerCountryCode = "BE"
documentTemplate = Replace(documentTemplate, "{customerSchemeId}", customerSchemeId)
'customerCompanyId = "04400580217"
documentTemplate = Replace(documentTemplate, "{customerCompanyId}", customerCompanyId)
'customerCompanyIdExtended = "BE04400580217"
documentTemplate = Replace(documentTemplate, "{customerCompanyIdExtended}", customerCompanyIdExtended)

'customerRegistrationName = "Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv"
documentTemplate = Replace(documentTemplate, "{customerRegistrationName}", customerRegistrationName)
'customerElectronicMail = "info@rv.be"
documentTemplate = Replace(documentTemplate, "{customerElectronicMail}", customerElectronicMail)

paymentMeansCode = "1"
documentTemplate = Replace(documentTemplate, "{paymentMeansCode}", paymentMeansCode)
'paymentID = "024010022505"
documentTemplate = Replace(documentTemplate, "{paymentId}", paymentID)
'supplierIBAN = "BE50001683378618"
documentTemplate = Replace(documentTemplate, "{supplierIBAN}", supplierIBAN)
'supplierBIC = "GEBABEBB"
documentTemplate = Replace(documentTemplate, "{supplierBIC}", supplierBIC)
'paymentTerms = "024010022505"
documentTemplate = Replace(documentTemplate, "{paymentTerms}", paymentTerms)

documentTemplate = Replace(documentTemplate, "{taxTotalAmount}", taxGlobalTotalAmount)
documentTemplate = Replace(documentTemplate, "{taxableAmount}", taxableGlobalTotalAmount)
documentTemplate = Replace(documentTemplate, "{taxExclusiveAmount}", taxableGlobalTotalAmount)

globalTotalAmount = Trim(Dec((Val(taxableGlobalTotalAmount) + Val(taxGlobalTotalAmount)), MASK_EUR))
documentTemplate = Replace(documentTemplate, "{taxInclusiveAmount}", globalTotalAmount)
documentTemplate = Replace(documentTemplate, "{payableAmount}", globalTotalAmount)

'Ktrl = ScrLeesBestandAlleTekst(invoiceTaxLineTemplate, PROGRAM_LOCATION + "xml-templates\ubl_be_3_0-invoicetaxline.xml")
'If Ktrl = 0 Then
'    MsgBox "Onverwachte situatie", vbCritical
'End If

'0%, 6%, 12%, 21%
For T = 0 To 3
    If BTWEuroBasis(T) <> 0 Then
        taxSubtotalLine = invoiceTaxLineTemplate
        subTaxableAmount = Trim(Dec((BTWEuroBasis(T)), MASK_EUR))
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxableAmount}", subTaxableAmount)
                
        subTaxAmount = Trim(Dec((BTWEuroBedrag(T)), MASK_EUR))
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxAmount}", subTaxAmount)
        subTaxCategory = Trim(Dec((T), "00"))
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxCategory}", subTaxCategory)
        subTaxPercent = Mid(fmarBoxText("002", "2", Trim(Str(T))), 4, 4)
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxPercent}", subTaxPercent)

        subTaxExclusiveAmount = Trim(Dec((BTWEuroBasis(T)), sy2))
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxExclusiveAmount}", subTaxExclusiveAmount)
        subTaxInclusiveAmount = Trim(Dec(BTWEuroBasis(T) + BTWEuroBedrag(T), MASK_EUR))
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxInclusiveAmount}", subTaxInclusiveAmount)
        subPayableAmount = taxInclusiveAmount
        taxSubtotalLine = Replace(taxSubtotalLine, "{payableAmount}", subPayableAmount)
        taxSubtotalLine = Replace(taxSubtotalLine, "{currency}", invoiceCurrency)
        taxSubtotalLine = Replace(taxSubtotalLine, "{supplierTaxScheme}", supplierTaxScheme)
        
        listTaxSubtotalLines = listTaxSubtotalLines + taxSubtotalLine
        If T = 3 Then
        Else
            listTaxSubtotalLines = listTaxSubtotalLines + vbCrLf
        End If
    End If
Next
'documentTemplate = Replace(documentTemplate, "<Vsoft>{taxsubtotals-template}</Vsoft>", listTaxSubtotalLines)

Ktrl = ScrLeesBestandAlleTekst(invoiceLineTemplate, PROGRAM_LOCATION + "xml-templates\ubl_be_3_0-invoiceline-no-vat.xml")
If Ktrl = 0 Then
    MsgBox "Onverwachte situatie", vbCritical
End If

ublTT = 0
For T = 0 To VerkoopDetail.ListCount - 1
    If T > 0 Then
        listInvoiceLines = listInvoiceLines + vbCrLf
    End If
    thisInvoiceLine = invoiceLineTemplate
    GridText = VerkoopDetail.List(T)
    If Right(GridText, 1) = "2" Then
        If T = VerkoopDetail.ListCount - 1 Then
            'Stop
        End If
    Else
        ublTT = ublTT + 1
        lineCounter = Trim(Str(ublTT))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineCounter+1}", lineCounter)
        lineDescription = Trim(Mid(GridText, 1, 40))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineDescription}", lineDescription)

        GridText = Right(GridText, Len(GridText) - 41)
        dVeldUblInfo(4) = Val(Left(GridText, 11))          'price exclusive without discount
        GridText = Right(GridText, Len(GridText) - 12)
        dVeldUblInfo(6) = Val(Left(GridText, 7))           'quantity
        GridText = Right(GridText, Len(GridText) - 8)
        dVeldUblInfo(7) = Val(Left(GridText, 12))          'total exclusive with discount
        GridText = Right(GridText, Len(GridText) - 13)
        dVeldUblInfo(2) = Val(Left(GridText, 6))           'maatstaf anders default 1.0
        GridText = Right(GridText, Len(GridText) - 7)
        VeldUblInfo(8) = Left(GridText, 1)                 '9 ??
        GridText = Right(GridText, Len(GridText) - 2)
        dVeldUblInfo(5) = Val(Left(GridText, 3))           'korting
        GridText = Right(GridText, Len(GridText) - 4)
        VeldUblInfo(9) = Left(GridText, 1)                 'vat id
        GridText = Right(GridText, Len(GridText) - 2)
        'If VeldUblInfo(9) = "6" Then
        '    VeldUblInfo(9) = "0"
        'End If
        dVeldUblInfo(3) = Val(Left(GridText, 7))           'number ledgeraccount
        GridText = Right(GridText, Len(GridText) - 8)
        VeldUblInfo(0) = Left(GridText, 13)                'product reference id or description
                
        linePriceAmount = Trim(Dec((dVeldUblInfo(7) / dVeldUblInfo(6)), MASK_EUR + "00"))
        thisInvoiceLine = Replace(thisInvoiceLine, "{linePriceAmount}", linePriceAmount)
                        
        lineQuantity = Trim(Dec(dVeldUblInfo(6), MASK_EUR))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineQuantity}", lineQuantity)
                
        lineExtentionAmount = Trim(Dec(dVeldUblInfo(7), MASK_EUR + "00"))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineExtensionAmount}", lineExtentionAmount)
        
        'tmpString = "701000;701000;701000"
        'lineAccountingcost = Split(tmpString, ";")
        'tmpString = "8.4000;83.8951;343.5600"
        'lineTaxAmountAccurate = Split(tmpString, ";")
        'tmpString = "1-701000;1-701000;1-701000"
        'lineSellerItemIdentification = Split(tmpString, ";")
        'tmpString = "03;03;03"
        'lineTaxCategory = Split(tmpString, ";")
        'tmpString = "21.00;21.00;21.00"
        'lineTaxPercentage = Split(tmpString, ";")

        lineAccountingcost = Trim(Str(dVeldUblInfo(3)))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineAccountingCost}", lineAccountingcost)
        
        If VeldUblInfo(9) = "6" Then
            VeldUblInfo(9) = "0"
        End If
        lineTaxCategory = Dec(Val(VeldUblInfo(9)), "00")
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxCategory}", lineTaxCategory)
        
        lineTaxPercentage = Mid(fmarBoxText("002", "2", VeldUblInfo(9)), 4, 4)
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxPercentage}", lineTaxPercentage)
                
        lineTaxAmountAccurate = Trim(Dec((dVeldUblInfo(7) * Val(lineTaxPercentage) / 100), MASK_EUR + "00"))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxAmountAccurate}", lineTaxAmountAccurate)
                
        lineSellerItemIdentification = Trim(VeldUblInfo(0))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineSellerItemIdentification}", lineSellerItemIdentification)
        
        thisInvoiceLine = Replace(thisInvoiceLine, "{currency}", invoiceCurrency)
        thisInvoiceLine = Replace(thisInvoiceLine, "{supplierTaxScheme}", supplierTaxScheme)
    
        listInvoiceLines = listInvoiceLines + thisInvoiceLine
    End If
Next
documentTemplate = Replace(documentTemplate, "<Vsoft>{invoicelines-template}</Vsoft>", listInvoiceLines)
Ktrl = ScrMaakTekstBestand(documentTemplate, LOCATION_COMPANYDATA + "peppol\out\" + invoiceNumber + ".xml")
If Ktrl = 0 Then
    MsgBox "Fout bij bewaren van " + invoiceNumber + ".xml", vbCritical
Else
    XmlInvoiceGenerateUBLBE30NoVat = True
End If

End Function

Function XmlInvoiceGenerateUBLV3(thisPdf As String) As Boolean

'UBL
'base64Str = ""
documentTemplate = ""
invoiceLineTemplate = ""
invoiceTaxLineTemplate = ""
taxSubtotalLine = ""
listTaxSubtotalLines = ""
thisInvoiceLine = ""
listInvoiceLines = ""
tmpString = ""
'SUPPLIER:
'invoiceNumber = "" '240100225
'invoiceCustomerNumber = "" ' 844
invoiceCurrency = "" 'EUR
'invoiceDate = "" '2024-01-31
'invoiceDueDate = "" '2024-03-01
'invoiceTypeCode = "" '380 (see https://docs.peppol.eu/poac/pint/pint/trn-invoice/codelist/UNCL1001-inv/)
'orderReference = "" 'ex. 024010022505 as of +++024/0100/22505+++
'supplierTaxScheme = "" 'ex. VAT
'supplierRegistrationId = "" 'ex. 0423100736
'supplierVatNumber = "" 'ex. 0423100736 (Titeca)
'supplierRegistrationName = "" 'Titeca Accountancy Merelbeke NV
'supplierStreetName = "" 'Fraterstraat 132
'supplierCityName = "" 'MERELBEKE
'supplierPostalZone = "" '9820
'supplierCountryCode = "" 'BE
'supplierCompanyId = "" 'BE0423100736
'supplierTelephone = "" '09 232 28 00
'CUSTOMER:
'customerVatNumber = "" '0440058217
'customerTaxScheme = "" 'ex. VAT
'if customer is VAT then else ?
'customerName = "" 'ex. Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv
'customerStreetName = "" 'Grote Baan 141
'customerCityName = "" 'HERDERSEM
'customerPostalZone = "" '9310
'customerCountryCode = "" 'BE
'customerCompanyId = "" 'BE04400580217
'customerRegistrationName = "" 'Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv
'customerElectronicMail = "" 'info@rv.be
'PAYMENTMEANS/TERMS
paymentMeansCode = "1" '1 see: https://docs.peppol.eu/poac/pint/pint/trn-invoice/codelist/UNCL4461/
'paymentID = "" '024010022505
'supplierIBAN = "" 'BE50001683378618
'supplierBIC = "" 'GEBABEBB
'paymentTerms = "" '024010022505
'TAXTOTAL
subTaxTotalAmount = "" '435.86
subTaxableAmount = "" '2075.50
subTaxAmount = "" '435.86
subTaxCategory = "" '03
subTaxPercent = "" '21.00
subTaxExclusiveAmount = ""
subTaxInclusiveAmount = ""
subPayableAmount = ""

'taxGlobalTotalAmount = ""
'taxableGlobalTotalAmount = ""
globalTotalAmount = ""

'LEGALMTOTAL
taxExclusiveAmount = "" '2075.50
taxInclusiveAmount = "" '2511.36
payableAmount = "" '2511.36

'LINES
lineCounter = ""
lineQuantity = ""
lineExtentionAmount = ""
lineAccountingcost = ""
lineTaxAmountAccurate = ""
lineDescription = ""
lineSellerItemIdentification = ""
lineTaxCategory = ""
lineTaxPercentage = ""
linePriceAmount = ""
vatGroupsCount = 0
ublTT = 0

XmlInvoiceGenerateUBLV3 = False
If Me.OptionPEPPOL_V3.Enabled = False Then
    Exit Function
End If

If Left(invoiceNumber, 2) = "V0" Then
    Ktrl = ScrLeesBestandAlleTekst(documentTemplate, PROGRAM_LOCATION + "xml-templates\peppol_bis_billing_ubl_v3-invoice.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
Else
    Ktrl = ScrLeesBestandAlleTekst(documentTemplate, PROGRAM_LOCATION + "xml-templates\peppol_bis_billing_ubl_v3-creditnote.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
End If

'invoiceNumber = "240100225"
documentTemplate = Replace(documentTemplate, "{documentNumber}", invoiceNumber)
'invoiceCustomerNumber = "844"
documentTemplate = Replace(documentTemplate, "{customerNumber}", invoiceCustomerNumber)
invoiceCurrency = "EUR"
documentTemplate = Replace(documentTemplate, "{currency}", invoiceCurrency)
    
'invoiceDate = "2024-01-31"
documentTemplate = Replace(documentTemplate, "{documentDate}", invoiceDate)
'invoiceDueDate = "2024-03-01"
If Left(invoiceNumber, 2) = "V0" Then
    documentTemplate = Replace(documentTemplate, "{documentDueDate}", invoiceDueDate)
End If

documentTemplate = Replace(documentTemplate, "{invoiceTypeCode}", invoiceTypeCode)
'orderReference = "024010022505" 'as of +++024/0100/22505+++
If orderMarReferences = "" Then
    documentTemplate = Replace(documentTemplate, "{orderReference}", invoiceNumber)
Else
    documentTemplate = Replace(documentTemplate, "{orderReference}", orderMarReferences)
End If

base64Str = GetBase64Str(thisPdf)
documentTemplate = Replace(documentTemplate, "{pdfFileName}", invoiceNumber + ".pdf")
documentTemplate = Replace(documentTemplate, "{pdfBase64}", base64Str)

documentTemplate = Replace(documentTemplate, "{supplierTaxScheme}", supplierTaxScheme)
'supplierVatNumber = "0423100736"
documentTemplate = Replace(documentTemplate, "{supplierVatNumber}", supplierVatNumber)
'tmpString = String99(READING, 46)
'If InStr(tmpString, "&") Then 'verbeteren voor XML bestand!!!
'    tmpString = Replace(tmpString, "&", "&amp;")
'End If
'supplierRegistrationId = "0423100736"
documentTemplate = Replace(documentTemplate, "{supplierRegistrationId}", supplierRegistrationId)
'supplierRegistrationName = "Titeca Accountancy Merelbeke NV"
documentTemplate = Replace(documentTemplate, "{supplierRegistrationName}", supplierRegistrationName)
'supplierStreetName = "Fraterstraat 132"
documentTemplate = Replace(documentTemplate, "{supplierStreetName}", supplierStreetName)
'supplierCityName = "MERELBEKE"
documentTemplate = Replace(documentTemplate, "{supplierCityName}", supplierCityName)
'supplierPostalZone = "9820"
documentTemplate = Replace(documentTemplate, "{supplierPostalZone}", supplierPostalZone)
'supplierCountryCode = "BE"
documentTemplate = Replace(documentTemplate, "{supplierCountryCode}", supplierCountryCode)
'supplierCompanyId = "0423100736"
documentTemplate = Replace(documentTemplate, "{supplierCompanyId}", supplierCompanyId)
'supplierCompanyIdExtended = "BE0423100736"
documentTemplate = Replace(documentTemplate, "{supplierCompanyIdExtended}", supplierCompanyIdExtended)

'supplierTelephone = "09 232 28 00"
documentTemplate = Replace(documentTemplate, "{supplierTelephone}", supplierTelephone)

'customerVatNumber = "0440058217"
documentTemplate = Replace(documentTemplate, "{customerVatNumber}", customerVatNumber)
documentTemplate = Replace(documentTemplate, "{customerTaxScheme}", customerTaxScheme)
'customerName = "Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv"
documentTemplate = Replace(documentTemplate, "{customerRegistrationName}", customerName)
'customerStreetName = "Grote Baan 141"
documentTemplate = Replace(documentTemplate, "{customerStreetName}", customerStreetName)
'customerCityName = "HERDERSEM"
documentTemplate = Replace(documentTemplate, "{customerCityName}", customerCityName)
'customerPostalZone = "9310"
documentTemplate = Replace(documentTemplate, "{customerPostalZone}", customerPostalZone)
'customerCountryCode = "BE"
documentTemplate = Replace(documentTemplate, "{customerCountryCode}", customerCountryCode)
'customerCountryCode = "BE"
documentTemplate = Replace(documentTemplate, "{customerSchemeId}", customerSchemeId)
'customerCompanyId = "04400580217"
documentTemplate = Replace(documentTemplate, "{customerCompanyId}", customerCompanyId)
'customerCompanyIdExtended = "BE04400580217"
documentTemplate = Replace(documentTemplate, "{customerCompanyIdExtended}", customerCompanyIdExtended)

'customerRegistrationName = "Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv"
documentTemplate = Replace(documentTemplate, "{customerRegistrationName}", customerRegistrationName)
'customerElectronicMail = "info@rv.be"
documentTemplate = Replace(documentTemplate, "{customerElectronicMail}", customerElectronicMail)

paymentMeansCode = "1"
documentTemplate = Replace(documentTemplate, "{paymentMeansCode}", paymentMeansCode)
'paymentID = "024010022505"
documentTemplate = Replace(documentTemplate, "{paymentId}", paymentID)
'supplierIBAN = "BE50001683378618"
documentTemplate = Replace(documentTemplate, "{supplierIBAN}", supplierIBAN)
'supplierBIC = "GEBABEBB"
documentTemplate = Replace(documentTemplate, "{supplierBIC}", supplierBIC)
'paymentTerms = "024010022505"
documentTemplate = Replace(documentTemplate, "{paymentTerms}", paymentTerms)

documentTemplate = Replace(documentTemplate, "{taxTotalAmount}", taxGlobalTotalAmount)
documentTemplate = Replace(documentTemplate, "{taxableAmount}", taxableGlobalTotalAmount)
documentTemplate = Replace(documentTemplate, "{taxExclusiveAmount}", taxableGlobalTotalAmount)

globalTotalAmount = Trim(Dec((Val(taxableGlobalTotalAmount) + Val(taxGlobalTotalAmount)), MASK_EUR))
documentTemplate = Replace(documentTemplate, "{taxInclusiveAmount}", globalTotalAmount)
documentTemplate = Replace(documentTemplate, "{payableAmount}", globalTotalAmount)

Ktrl = ScrLeesBestandAlleTekst(invoiceTaxLineTemplate, PROGRAM_LOCATION + "xml-templates\peppol_bis_billing_ubl_v3-invoicetaxline.xml")
If Ktrl = 0 Then
    MsgBox "Onverwachte situatie", vbCritical
End If

'0%, 6%, 12%, 21%
For T = 0 To 3
    If BTWEuroBasis(T) <> 0 Then
        taxSubtotalLine = invoiceTaxLineTemplate
        subTaxableAmount = Trim(Dec((BTWEuroBasis(T)), MASK_EUR))
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxableAmount}", subTaxableAmount)
                
        If T = 0 Then
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxAmount}", "0")
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxIndex}", "Z")
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxCategory}", "00")
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxPercent}", "0")
        Else
            subTaxAmount = Trim(Dec((BTWEuroBedrag(T)), MASK_EUR))
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxAmount}", subTaxAmount)
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxIndex}", "S")
            subTaxCategory = Trim(Dec((T), "00"))
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxCategory}", subTaxCategory)
            subTaxPercent = Mid(fmarBoxText("002", "2", Trim(Str(T))), 4, 4)
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxPercent}", subTaxPercent)
        End If

        subTaxExclusiveAmount = Trim(Dec((BTWEuroBasis(T)), sy2))
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxExclusiveAmount}", subTaxExclusiveAmount)
        subTaxInclusiveAmount = Trim(Dec(BTWEuroBasis(T) + BTWEuroBedrag(T), MASK_EUR))
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxInclusiveAmount}", subTaxInclusiveAmount)
        subPayableAmount = taxInclusiveAmount
        taxSubtotalLine = Replace(taxSubtotalLine, "{payableAmount}", subPayableAmount)
        taxSubtotalLine = Replace(taxSubtotalLine, "{currency}", invoiceCurrency)
        taxSubtotalLine = Replace(taxSubtotalLine, "{supplierTaxScheme}", supplierTaxScheme)
        
        listTaxSubtotalLines = listTaxSubtotalLines + taxSubtotalLine
        If T = 3 Then
        Else
            listTaxSubtotalLines = listTaxSubtotalLines + vbCrLf
        End If
    End If
Next
documentTemplate = Replace(documentTemplate, "<Vsoft>{taxsubtotals-template}</Vsoft>", listTaxSubtotalLines)

If Left(invoiceNumber, 2) = "V0" Then
    Ktrl = ScrLeesBestandAlleTekst(invoiceLineTemplate, PROGRAM_LOCATION + "xml-templates\peppol_bis_billing_ubl_v3-invoiceline.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
Else
    Ktrl = ScrLeesBestandAlleTekst(invoiceLineTemplate, PROGRAM_LOCATION + "xml-templates\peppol_bis_billing_ubl_v3-creditnoteline.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
End If

ublTT = 0
For T = 0 To VerkoopDetail.ListCount - 1
    If T > 0 Then
        listInvoiceLines = listInvoiceLines + vbCrLf
    End If
    thisInvoiceLine = invoiceLineTemplate
    GridText = VerkoopDetail.List(T)
    If Right(GridText, 1) = "2" Then
        If T = VerkoopDetail.ListCount - 1 Then
            'Stop
        End If
    Else
        ublTT = ublTT + 1
        lineCounter = Trim(Str(ublTT))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineCounter+1}", lineCounter)
        lineDescription = Trim(Mid(GridText, 1, 40))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineDescription}", lineDescription)

        GridText = Right(GridText, Len(GridText) - 41)
        dVeldUblInfo(4) = Val(Left(GridText, 11))          'price exclusive without discount
        GridText = Right(GridText, Len(GridText) - 12)
        dVeldUblInfo(6) = Val(Left(GridText, 7))           'quantity
        GridText = Right(GridText, Len(GridText) - 8)
        dVeldUblInfo(7) = Val(Left(GridText, 12))          'total exclusive with discount
        GridText = Right(GridText, Len(GridText) - 13)
        dVeldUblInfo(2) = Val(Left(GridText, 6))           'maatstaf anders default 1.0
        GridText = Right(GridText, Len(GridText) - 7)
        VeldUblInfo(8) = Left(GridText, 1)                 '9 ??
        GridText = Right(GridText, Len(GridText) - 2)
        dVeldUblInfo(5) = Val(Left(GridText, 3))           'korting
        GridText = Right(GridText, Len(GridText) - 4)
        VeldUblInfo(9) = Left(GridText, 1)                 'vat id
        GridText = Right(GridText, Len(GridText) - 2)
        'If VeldUblInfo(9) = "6" Then
        '    VeldUblInfo(9) = "0"
        'End If
        dVeldUblInfo(3) = Val(Left(GridText, 7))           'number ledgeraccount
        GridText = Right(GridText, Len(GridText) - 8)
        VeldUblInfo(0) = Left(GridText, 13)                'product reference id or description
                
        linePriceAmount = Trim(Dec((dVeldUblInfo(7) / dVeldUblInfo(6)), MASK_EUR))
        thisInvoiceLine = Replace(thisInvoiceLine, "{linePriceAmount}", linePriceAmount)
                        
        lineQuantity = Trim(Dec(dVeldUblInfo(6), MASK_EUR))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineQuantity}", lineQuantity)
                
        lineExtentionAmount = Trim(Dec(dVeldUblInfo(7), MASK_EUR))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineExtensionAmount}", lineExtentionAmount)
        
        'tmpString = "701000;701000;701000"
        'lineAccountingcost = Split(tmpString, ";")
        'tmpString = "8.4000;83.8951;343.5600"
        'lineTaxAmountAccurate = Split(tmpString, ";")
        'tmpString = "1-701000;1-701000;1-701000"
        'lineSellerItemIdentification = Split(tmpString, ";")
        'tmpString = "03;03;03"
        'lineTaxCategory = Split(tmpString, ";")
        'tmpString = "21.00;21.00;21.00"
        'lineTaxPercentage = Split(tmpString, ";")

        lineAccountingcost = Trim(Str(dVeldUblInfo(3)))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineAccountingCost}", lineAccountingcost)
        
        If VeldUblInfo(9) = "6" Then
            VeldUblInfo(9) = "0"
            thisInvoiceLine = Replace(thisInvoiceLine, "{taxIndex}", "Z")
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxCategory}", "00")
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxPercentage}", "0")
            
            lineTaxAmountAccurate = "0.0000"
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxAmountAccurate}", lineTaxAmountAccurate)
        Else
            thisInvoiceLine = Replace(thisInvoiceLine, "{taxIndex}", "S")
            
            lineTaxCategory = Dec(Val(VeldUblInfo(9)), "00")
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxCategory}", lineTaxCategory)
        
            lineTaxPercentage = Mid(fmarBoxText("002", "2", VeldUblInfo(9)), 4, 4)
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxPercentage}", lineTaxPercentage)
        
            lineTaxAmountAccurate = Trim(Dec((dVeldUblInfo(7) * Val(lineTaxPercentage) / 100), MASK_EUR))
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxAmountAccurate}", lineTaxAmountAccurate)
        End If
                
        lineSellerItemIdentification = Trim(VeldUblInfo(0))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineSellerItemIdentification}", lineSellerItemIdentification)
        
        thisInvoiceLine = Replace(thisInvoiceLine, "{currency}", invoiceCurrency)
        thisInvoiceLine = Replace(thisInvoiceLine, "{supplierTaxScheme}", supplierTaxScheme)
    
        listInvoiceLines = listInvoiceLines + thisInvoiceLine
    End If
Next

If Left(invoiceNumber, 2) = "V0" Then
    documentTemplate = Replace(documentTemplate, "<Vsoft>{invoicelines-template}</Vsoft>", listInvoiceLines)
Else
    documentTemplate = Replace(documentTemplate, "<Vsoft>{creditnotelines-template}</Vsoft>", listInvoiceLines)
End If
Ktrl = ScrMaakTekstBestand(documentTemplate, LOCATION_COMPANYDATA + "peppol\out\" + invoiceNumber + ".xml")
If Ktrl = 0 Then
    MsgBox "Fout bij bewaren van " + invoiceNumber + ".xml", vbCritical
Else
    XmlInvoiceGenerateUBLV3 = True
End If

End Function

Private Sub Afsluiten_Click()
Dim A As String
Dim aa As String * 30
Dim BonnenString As String
Dim dTTwb               As Double
Dim DummySleutel As String
Dim BestondReeds    As Integer
Dim T As Integer
Dim tempBDos As Integer

' REFRESH CLIENT FIRST!!??
'XLogKey = Trim(Left(KlantInfo.Caption, 12))
'bGet TABLE_CUSTOMERS, 0, XLogKey
'If Ktrl Then
'    MsgBox "error"
'Else
'    RecordToVeld TABLE_CUSTOMERS
'End If

printAndSave = True

If Vr = 11 Or Vr = 13 Then
    If dokumentSleutel <> SleutelDok(Vr) Then
    Msg = dokumentSleutel + " <> " + SleutelDok(Vr) + vbCrLf + vbCrLf
    Msg = Msg + "MOGELIJKHEID 1: Teller is identiek, boekjaar is hoger/lager." + vbCrLf
    Msg = Msg + "U hebt dus het actief boekjaar of periode gewijzigd tijdens de aanmaak van dit dokument.  Probeer nogmaals NA KONTROLE." + vbCrLf + vbCrLf
    Msg = Msg + "MOGELIJKHEID 2: Boekjaar is identiek, teller is hoger/lager." + vbCrLf
    Msg = Msg + "Een andere gebruiker heeft ondertussen een dokument verwerkt." + vbCrLf + vbCrLf
    Msg = Msg + "Kontroleer eerst eens vooraleer de boeking nogmaals uit te voeren a.u.b. !!!"
    MsgBox Msg
    dokumentSleutel = SleutelDok(Vr)
    TekstInfo3 = dokumentSleutel
    DirekteVerkoop.Caption = vSet("Verkoopverrichting", 28) + "(" + dokumentSleutel + ")"
    Exit Sub
    End If
End If

If dokumentType = "15" And Annuleren.Enabled = True Then
    If Not DATE_CHECK((TekstInfo0.Text), PERIODAS_TEXT) Then
        Beep
        TekstInfo0.SetFocus
        Exit Sub
    End If
End If

Msg = "Document administratief afhandelen" + vbCrLf + vbCrLf
Msg = Msg + "Bent U zeker ?"
Ktrl = MsgBox(Msg$, vbYesNo + vbDefaultButton2 + vbQuestion, "Document : " + dokumentSleutel)
Select Case Ktrl
    Case vbYes
    Case Else
        Exit Sub
End Select

If SteedsDrukken Then
    Afdrukken
End If
      
If bhEuro And cmdSwitch.Caption = "Ingave in EUR" Then
    'ok VOOR EUR boekhouding
ElseIf Not bhEuro And cmdSwitch.Caption = "Ingave in BEF" Then
    'ok VOOR BEF boekhouding
Else
    cmdSwitch_Click
    DoEvents
End If
TotaalBTW = 0
If Medekontraktant.Value Or VerkoopFLG Then
    BTWBedrag(1) = 0: BTWEuroBedrag(1) = 0
    BTWBedrag(2) = 0: BTWEuroBedrag(2) = 0
    BTWBedrag(3) = 0: BTWEuroBedrag(3) = 0
Else
    For T = 0 To 3
        If bhEuro Then
            TotaalBTW = TotaalBTW + BTWEuroBedrag(T)
        Else
            TotaalBTW = TotaalBTW + BTWBedrag(T)
        End If
    Next
End If

If VerkoopFLG Then
    If bhEuro Then
        TotaalUitvoer = CDbl(LblEx2Btw)
    Else
        TotaalUitvoer = CDbl(LblExBtw)
    End If
    For T = 0 To 3
        BTWEuroBasis(T) = 0
        BTWBasis(T) = 0
    Next
Else
    TotaalUitvoer = 0
End If

Dim sellersDoc As String

Err = 0
On Local Error Resume Next
If Annuleren.Enabled = False Then
    GoTo Nextdokument
ElseIf dokumentType = "15" Then
    If frmDokHistoriek.lstDokHistoriek.ListCount <> 0 Then
        BonnenString = ""
        For T = 0 To frmDokHistoriek.lstDokHistoriek.ListCount - 1
            BonnenString = BonnenString + frmDokHistoriek.lstDokHistoriek.List(T) + ";"
        Next
    End If
    frmDokHistoriek.Hide
    Msg = "Factuur of Creditnota !" + vbCrLf + vbCrLf + "Hierna wordt de boekhouding bijgewerkt.  Bent U zeker ?"
    Ktrl = MsgBox(Msg$, 292, "Boekhouding bijwerken")
    If Ktrl = 6 Then
        Ktrl = bOpen(TABLE_LEDGERACCOUNTS)
        Ktrl = bOpen(TABLE_JOURNAL)
        Ktrl = bOpen(TABLE_INVOICES)
        Ktrl = bOpen(TABLE_PRODUCTS)
        bBegin
        If WegBoekFout() Then
            bAbort
            'look for documents to print or send with peppol and remove them
            On Error Resume Next
            Err = 0
            sellersDoc = Dir(LOCATION_COMPANYDATA + "peppol\out\" + invoiceNumber + ".xml")
            If sellersDoc <> "" Then
                Kill LOCATION_COMPANYDATA + "peppol\out\" + invoiceNumber + ".xml"
                MsgBox LOCATION_COMPANYDATA + "peppol\out\" + invoiceNumber + ".xml is verwijderd en boeking genegeerd.", vbInformation
            End If
            
            On Error Resume Next
            Err = 0
            'bureaublad?
            sellersDoc = Dir(desktopLocatie & "\" & invoiceNumber & ".xml")
                        
            On Error Resume Next
            Err = 0
            If sellersDoc <> "" Then
                Kill desktopLocatie & "\" & invoiceNumber & ".xml"
                MsgBox desktopLocatie & "\" & invoiceNumber & ".xml is eveneens verwijderd.", vbInformation
            End If
            Exit Sub
        Else
            bEnd
            If VerkoopFLG = 1 Then
                Fl = TABLE_CUSTOMERS
                aIndex = 29
                dTTwb = Val(vBibTekst(TABLE_INVOICES, "#v060 #")) + Val(vBibTekst(TABLE_INVOICES, "#v062 #")) + Val(vBibTekst(TABLE_INVOICES, "#v089 #"))
                GridText = Dec$(dTTwb, MASK_SY(0)) + vbTab
                Intrastat.Show 1
            End If
        End If
    Else
        On Error Resume Next
        Err = 0
        sellersDoc = Dir(LOCATION_COMPANYDATA + "peppol\out\" + invoiceNumber + ".xml")
        If sellersDoc <> "" Then
            Kill LOCATION_COMPANYDATA + "peppol\out\" + invoiceNumber + ".xml"
            MsgBox LOCATION_COMPANYDATA + "peppol\out\" + invoiceNumber + ".xml is verwijderd en boeking genegeerd.", vbInformation
        End If
            
        On Error Resume Next
        Err = 0
        'bureaublad?
        sellersDoc = Dir(desktopLocatie & "\" & invoiceNumber & ".xml")
                        
        On Error Resume Next
        Err = 0
        If sellersDoc <> "" Then
            Kill desktopLocatie & "\" & invoiceNumber & ".xml"
            MsgBox desktopLocatie & "\" & invoiceNumber & ".xml is eveneens verwijderd.", vbInformation
        End If
        
        On Error Resume Next
        Err = 0
        'manueel?
        sellersDoc = Dir(locPOSTVAKIN & "\" & idPdfForUbl)
                        
        On Error Resume Next
        Err = 0
        If sellersDoc <> "" Then
            Kill locPOSTVAKIN & "\" & idPdfForUbl
            MsgBox locPOSTVAKIN & "\" & idPdfForUbl & " is eveneens verwijderd.", vbInformation
        End If
        
        
        Exit Sub
    End If
End If
If Err Then MsgBox Error

bGet TABLE_VARIOUS, 1, vSet(dokumentType + dokumentSleutel, 20)
If Ktrl Then
    BestondReeds = False
Else
    BestondReeds = True
    bDelete TABLE_VARIOUS
    Do
        bNext TABLE_VARIOUS
        If Ktrl Or vSet(KEY_BUF(TABLE_VARIOUS), 20) <> vSet(dokumentType + dokumentSleutel, 20) Then
            Exit Do
        Else
            bDelete TABLE_VARIOUS
        End If
    Loop
End If

tempBDos = 0
JumpNextBlok:
TLB_RECORD(TABLE_VARIOUS) = ""
vBib TABLE_VARIOUS, "K" + RV(rsKlant, "A110"), "v004"
vBib TABLE_VARIOUS, dokumentType + dokumentSleutel, "v005"
vBib TABLE_VARIOUS, RV(rsKlant, "A110"), "A110"
vBib TABLE_VARIOUS, dokumentSleutel, "v033"
vBib TABLE_VARIOUS, Format(Medekontraktant.Value), "v135"
vBib TABLE_VARIOUS, DATE_KEY((TekstInfo0.Text)), "v035"
vBib TABLE_VARIOUS, DATE_KEY((TekstInfo1.Text)), "v036"
vBib TABLE_VARIOUS, (TekstInfo2.Text), "v136"
If bhEuro Then
    vBib TABLE_VARIOUS, (LblEx2Btw), "v137"
    vBib TABLE_VARIOUS, (LblIn2Btw), "v138"
    vBib TABLE_VARIOUS, "EUR", "vEUR"
Else
    vBib TABLE_VARIOUS, (LblExBtw), "v137"
    vBib TABLE_VARIOUS, (LblInBtw), "v138"
End If
vBib TABLE_VARIOUS, (TekstInfo5.Text), "v040"
vBib TABLE_VARIOUS, Right(DagKoers.Caption, 3), "vs03"
If dokumentType = "15" Then
    If BonnenString <> "" Then
        vBib TABLE_VARIOUS, BonnenString, "v147"
    End If
End If

For T = 0 To 7
    If tempBDos + T > VerkoopDetail.ListCount - 1 Then
        Exit For
    Else
        VerkoopDetail.ListIndex = tempBDos + T
        vBib TABLE_VARIOUS, (VerkoopDetail.Text), "v" + Format(139 + T, "000")
    End If
Next
bInsert TABLE_VARIOUS, 1
If Ktrl Then
    MsgBox "stop dokument tekst wegschrijven!"
End If
If tempBDos + T > VerkoopDetail.ListCount - 1 Then
Else
    tempBDos = tempBDos + 8
    GoTo JumpNextBlok
End If
If BestondReeds = True Then
    GoTo Nextdokument
End If

SS99 Right(dokumentSleutel, 5), Vr
If BonnenString <> "" Then
    Do While BonnenString <> ""
        dokumentHistoriek = Left(BonnenString, 11)
        BonnenString = Right(BonnenString, Len(BonnenString) - 12)

    Select Case Left(dokumentHistoriek, 2)
        Case "V0", "V1"
            MsgBox "kanniet : onlogisch dokument", vbCritical, dokumentHistoriek

        Case "B0"
            DummySleutel = "14"
        Case "F0"
            DummySleutel = "13"
    End Select
    
    bGet TABLE_VARIOUS, 1, vSet(DummySleutel + dokumentHistoriek, 20)
    If Ktrl Then
        MsgBox "Onverwachte stop dokument " + dokumentHistoriek
    Else
        Do
            RecordToVeld TABLE_VARIOUS
            vBib TABLE_VARIOUS, dokumentSleutel, "v147"
            bUpdate TABLE_VARIOUS, 1
            If Ktrl Then
                MsgBox "Stop"
            End If
            bNext TABLE_VARIOUS
            If Ktrl Or Trim$(KEY_BUF(TABLE_VARIOUS)) <> DummySleutel + dokumentHistoriek Then
                Exit Do
            End If
        Loop
    End If
    Loop
ElseIf Trim$(dokumentHistoriek) <> "" Then
    Select Case Left(dokumentHistoriek, 2)
        Case "V0", "V1"
            DummySleutel = "15"
        Case "B0"
            DummySleutel = "14"
        Case "F0"
            DummySleutel = "13"
    End Select
    bGet TABLE_VARIOUS, 1, vSet(DummySleutel + dokumentHistoriek, 20)
    If Ktrl Then
        MsgBox "Onverwachte stop dokument " + dokumentHistoriek
    Else
        RecordToVeld TABLE_VARIOUS
        vBib TABLE_VARIOUS, dokumentSleutel, "v147"
        bUpdate TABLE_VARIOUS, 1
        If Ktrl Then
            MsgBox "Stop"
        End If
    End If
End If

Nextdokument:
dokumentSleutel = SleutelDok(Vr)
TekstInfo3 = dokumentSleutel
DirekteVerkoop.Caption = Left(DirekteVerkoop.Caption, 28) + "(" + dokumentSleutel + ")"
SchoonVegen_Click
KlantAktiveren.SetFocus

End Sub


Private Sub Annuleren_Click()

If VerkoopDetail.ListCount Then
          Msg = "Huidige inbreng negeren en venster sluiten." + vbCrLf + vbCrLf
    Msg = Msg + "Bent U zeker ?"
    Ktrl = MsgBox(Msg, 292, "Verkoopverrichtingen verlaten")
    If Ktrl = 6 Then
    Else
        Exit Sub
    End If
End If
Unload Me

End Sub


Private Sub ButtonInfoSupported_Click()

    MsgBox rsKlant("v407"), vbInformation, "Ondersteunde documenten"
    
End Sub

Private Sub cbCheckTools_Click()
    
    Load ValidatingTool

    ValidatingTool.Caption = "Klanten Peppol Tools"
    ValidatingTool.tbCompanyNumber = customerCompanyId
    ValidatingTool.tbVatNumber = customerCountryCode + customerCompanyId
    ValidatingTool.tbPeppolID = "0208:" + customerCompanyId
    ValidatingTool.Show 1
    If Trim(ValidatingTool.TextBoxSupportedDocuments) = "" Then
    Else
        If rsKlant("v407") = ValidatingTool.TextBoxSupportedDocuments.Text Then
            MsgBox "Er zijn geen wijzigingen aan ondersteunde documenten", vbInformation
        Else
            Msg = "Ondersteunde Peppol verrichtingen bijwerken voor deze klant" & vbCrLf & vbCrLf
            Msg = Msg + "Bent U zeker"
            KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2 + vbQuestion)
            If KtrlBox = vbYes Then
                rsKlant("v407") = ValidatingTool.TextBoxSupportedDocuments.Text
                rsKlant.Update
                
                MsgBox "Ondersteunde verrichtingen bijgewerkt. Herstart de verkoopprocedure", vbInformation
                Schoon
            End If
        End If
    End If
    Unload ValidatingTool
    
End Sub


Private Sub cbFactureren_Click()

    MsgBox lvDetail.ListItems.Count - 1
    
End Sub

Private Sub cbLayOudPDF_Click()

   LayOutpdfDokument.Show

End Sub

Private Sub cbMAPIMAIL_Click()

    Screen.MousePointer = vbHourglass
    'frmMAPIMailCTRL.Show

End Sub

Private Sub cbMonitortUBL_Click()
    
    Dim counterToSend As Integer
    Dim counterSent As Integer
    
    Dim flHier As Integer
    Dim ublFileUrl As String
    
    Screen.MousePointer = vbHourglass
    On Local Error Resume Next
    
    FormB2BVMonitor.Hide
    
    FormB2BVMonitor.mfgToSend.Clear
    FormB2BVMonitor.mfgToSend.Cols = 6
    FormB2BVMonitor.mfgToSend.Rows = 2
                        
    FormB2BVMonitor.mfgToSend.TextMatrix(0, 0) = "Klant"
    FormB2BVMonitor.mfgToSend.TextMatrix(0, 1) = "Ondernemingsnummer"
    FormB2BVMonitor.mfgToSend.TextMatrix(0, 2) = "Type Doc."
    FormB2BVMonitor.mfgToSend.TextMatrix(0, 3) = "Datum Doc"
    FormB2BVMonitor.mfgToSend.TextMatrix(0, 4) = "Vervaldag"
    FormB2BVMonitor.mfgToSend.TextMatrix(0, 5) = "Bestand"
    
    FormB2BVMonitor.mfgToSend.ColWidth(0) = 2130
    FormB2BVMonitor.mfgToSend.ColWidth(1) = 1710
    FormB2BVMonitor.mfgToSend.ColWidth(2) = 1110
    FormB2BVMonitor.mfgToSend.ColWidth(3) = 990
    FormB2BVMonitor.mfgToSend.ColWidth(4) = 990
    FormB2BVMonitor.mfgToSend.ColWidth(5) = 2130
    
    FormB2BVMonitor.mfgSent.Clear
    FormB2BVMonitor.mfgSent.Cols = 6
    FormB2BVMonitor.mfgSent.Rows = 2
    
    FormB2BVMonitor.mfgSent.TextMatrix(0, 0) = "Klant"
    FormB2BVMonitor.mfgSent.TextMatrix(0, 1) = "Type Doc."
    FormB2BVMonitor.mfgSent.TextMatrix(0, 2) = "Datum Doc"
    FormB2BVMonitor.mfgSent.TextMatrix(0, 3) = "Vervaldag"
    FormB2BVMonitor.mfgSent.TextMatrix(0, 4) = "Status"
    FormB2BVMonitor.mfgSent.TextMatrix(0, 5) = "Bestand"
    
    FormB2BVMonitor.mfgSent.ColWidth(0) = 2130
    FormB2BVMonitor.mfgSent.ColWidth(1) = 1110
    FormB2BVMonitor.mfgSent.ColWidth(2) = 990
    FormB2BVMonitor.mfgSent.ColWidth(3) = 990
    FormB2BVMonitor.mfgSent.ColWidth(4) = 1710
    FormB2BVMonitor.mfgSent.ColWidth(5) = 2130
        
    Dim c As Long
    For c = 0 To 5
        FormB2BVMonitor.mfgToSend.ColAlignment(c) = flexAlignLeftCenter
        FormB2BVMonitor.mfgSent.ColAlignment(c) = flexAlignLeftCenter
    Next
        
    Dim sPath As String
    Dim sFile As String
    
    sPath = LOCATION_COMPANYDATA & "peppol\out\"
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
    sFile = Dir$(sPath & "*.xml")          ' first match
    Do While sFile <> ""
        On Local Error Resume Next
        
        Dim documentCode As String
        Dim listItemToAdd As String
        
        Dim result As String
        
        result = CheckDocument(sFile)
                
        If Mid(sFile, 1, 2) = "V0" Then
            documentCode = "Factuur"
        ElseIf Mid(sFile, 1, 2) = "V1" Then
            documentCode = "Creditnota"
        Else
            documentCode = "??"
        End If
            
        Dim dateDocument As String
        Dim entry As String
        
        dateDocument = DATE_TEXT(rsSellerUBL("v035"))
            
        Dim dateExpiringDocument As String
        dateExpiringDocument = DATE_TEXT(rsSellerUBL("v036"))
           
        If Mid(result, 1, 12) = "Verzonden   " Then
            counterSent = counterSent + 1
            entry = rsSellerUBL("A100") & vbTab & documentCode & vbTab & dateDocument & vbTab & dateExpiringDocument & vbTab & Mid(result, 13) & vbTab & sFile
            FormB2BVMonitor.mfgSent.AddItem entry, counterSent
        Else
            counterToSend = counterToSend + 1
            entry = rsSellerUBL("A100") & vbTab & rsSellerUBL("v404") & vbTab & documentCode & vbTab & dateDocument & vbTab & dateExpiringDocument & vbTab & sFile
            FormB2BVMonitor.mfgToSend.AddItem entry, counterToSend
        End If
        sFile = Dir$()                      ' subsequent matches
    Loop
    Screen.MousePointer = vbNormal
    XLogKey = ""
    FormB2BVMonitor.Show 1

End Sub


Private Sub cbSelect_Click()

    Dim T As Integer
    
    GridText = ""
    For T = 1 To lvDetail.ListItems.Count
        lvDetail.ListItems(T).Checked = Not lvDetail.ListItems(T).Checked
    Next
        
End Sub



Private Sub cbUncl1001_Click()
    
    invoiceTypeCode = Mid(Me.cbUncl1001.Text, 1, InStr(Me.cbUncl1001.Text, ":") - 1)
    
End Sub


Private Sub cbUncl1001Toggle_Click()

Me.cbUncl1001.Enabled = Not Me.cbUncl1001.Enabled


End Sub

Private Sub chkBTWBouw_Click()
Dim Telhier As Integer
Dim StringHier As String

If chkBTWBouw.Value = vbChecked Then
    Msg = "Alle verkooplijnen omwerken tot BTW code '1'" + vbCr + vbCr
    Msg = Msg + "Bent U zeker"
    KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
    If KtrlBox = vbYes Then
        Do While Telhier <= VerkoopDetail.ListCount - 1
            If Right(DirekteVerkoop.VerkoopDetail.List(Telhier), 1) = "2" Then
            Else
                If Mid(DirekteVerkoop.VerkoopDetail.List(Telhier), 88, 1) <> "1" Then
                    StringHier = DirekteVerkoop.VerkoopDetail.List(Telhier)
                    Mid(StringHier, 88, 1) = "1"
                    VerkoopDetail.RemoveItem Telhier
                    VerkoopDetail.AddItem StringHier, Telhier
                End If
            End If
            Telhier = Telhier + 1
        Loop
        RefreshBTW
        VAT_BOBTHEBUILDERS = True
    Else
        chkBTWBouw.Value = 0
        VAT_BOBTHEBUILDERS = False
        Exit Sub
    End If
Else
    VAT_BOBTHEBUILDERS = False
End If

End Sub

Private Function Afdrukken()
Dim A As String
Dim refString As String
Dim refID As String

Dim aa As String * 30
Dim dTTwb               As Double
Dim DummySleutel As String
Dim BestondReeds    As Integer
Dim T As Integer

' REFRESH CLIENT FIRST!!
XLogKey = Trim(Left(KlantInfo.Caption, 12))
bGet TABLE_CUSTOMERS, 0, XLogKey
If Ktrl Then
    MsgBox "error"
Else
    RecordToVeld TABLE_CUSTOMERS
End If

If dokumentType = "15" And Annuleren.Enabled = True Then
    If Not DATE_CHECK((TekstInfo0.Text), PERIODAS_TEXT) Then
        Beep
        BJPERDAT.WindowState = 0
        BJPERDAT.PeriodeBoekjaar.SetFocus
        Exit Function
    End If
End If

orderMarReferences = ""
If dokumentType = "15" Then
    If frmDokHistoriek.lstDokHistoriek.ListCount <> 0 Then
        Msg = ""
        refID = "RF:"
        For T = 0 To frmDokHistoriek.lstDokHistoriek.ListCount - 1
            Msg = Msg + frmDokHistoriek.lstDokHistoriek.List(T) + " "
        Next
        KtrlBox = MsgBox("Bons als referte opnemen." & vbCr & vbCr & refID & Msg, vbYesNo + vbQuestion + vbDefaultButton1)
        If KtrlBox = vbYes Then
            refString = String(75, " ") + "|2"
            VerkoopDetail.AddItem refString
            Do While Msg <> ""
                refString = String(75, " ") + "|2"
                Mid(refString, 1) = refID + Left(Msg, 72)
                If Len(Msg) > 72 Then
                    Msg = Mid(Msg, 73)
                Else
                    Msg = ""
                End If
                VerkoopDetail.AddItem refString
            Loop
            orderMarReferences = Trim(frmDokHistoriek.lstDokHistoriek.List(0))
        Else
            orderMarReferences = ""
        End If
    End If
End If
    
TotaalBTW = 0
If Medekontraktant.Value Or VerkoopFLG Then
    BTWBedrag(1) = 0: BTWEuroBedrag(1) = 0
    BTWBedrag(2) = 0: BTWEuroBedrag(2) = 0
    BTWBedrag(3) = 0: BTWEuroBedrag(3) = 0
Else
    For T = 0 To 3
        If InStr(cmdSwitch.Caption, "EUR") Then
            TotaalBTW = TotaalBTW + BTWEuroBedrag(T)
        Else
            TotaalBTW = TotaalBTW + BTWBedrag(T)
        End If
    Next
End If

If VerkoopFLG Then
    If sMuntKlant = "EUR" Then
        TotaalUitvoer = CDbl(LblEx2Btw)
    Else
        TotaalUitvoer = CDbl(LblExBtw)
    End If
    For T = 0 To 3
        BTWBasis(T) = 0
        BTWEuroBasis(T) = 0
    Next
Else
    TotaalUitvoer = 0
End If

Dim mailAddressV224 As String
Dim nameAddressV224 As String

Dim mailAddressv002 As String
Dim nameAddressv002 As String

Dim voorkeurg101 As String
Dim idClient As String
Dim idDocument As String
Dim idMailAddress As String

Dim VMAIL_TO As Long
VMAIL_TO = 1


If Mim.Report.IsOpen = True Then
    Mim.Report.CloseDoc
End If
        
With Mim.Report
    .OpenDoc
    .Author = Trim(String99(READING, 46))
    .GUILanguage = 3 'Nederlands
    .Title = "VerkoopDocument"
End With
pdfDrukAf
        
idClient = Trim(rsMAR(TABLE_CUSTOMERS)("A110"))
If Not IsNull(rsMAR(TABLE_CUSTOMERS)("g101")) Then voorkeurg101 = Trim(rsMAR(TABLE_CUSTOMERS)("g101"))
idDocument = dokumentSleutel
tmpString = DATE_KEY(TekstInfo0.Text)
invoiceDate = Mid(tmpString, 1, 4) + "-" + Mid(tmpString, 5, 2) + "-" + Mid(tmpString, 7, 2)
tmpString = DATE_KEY(TekstInfo1.Text)
invoiceDueDate = Mid(tmpString, 1, 4) + "-" + Mid(tmpString, 5, 2) + "-" + Mid(tmpString, 7, 2)
    
'idPdfForUbl = "{" + idDocument + "}-{" + DATE_KEY(TekstInfo1.Text) + "}-{" + idClient + "}.pdf"
idPdfForUbl = "{" + idDocument + "}-{" + idClient + "}.pdf"
    
'always a copy first in io map even when no mail selected
If Not IsNull(rsMAR(TABLE_CUSTOMERS)("V224")) Then
    mailAddressV224 = Trim(rsMAR(TABLE_CUSTOMERS)("V224"))
    nameAddressV224 = Trim(Trim(rsMAR(TABLE_CUSTOMERS)("A100")) & " " & Trim(rsMAR(TABLE_CUSTOMERS)("A101")))
    If String99(READING, 306) = "2" Then
        Mim.Report.WriteDoc (locPOSTVAKIN & "\" & idPdfForUbl)
    End If
End If
            
If reprintOnly = False Then
    If Trim(customerVatNumber) = "" Then
    ElseIf VerkoopOptie(0).Value = True Then
        Mim.Report.WriteDoc (locPOSTVAKIN & "\" & idPdfForUbl)
        Dim peppolSuccess As Boolean
        If Me.OptionUBL_BE_3_0.Visible And Me.OptionPEPPOL_V3.Visible = True Then
            customerPrefersEmail = False
            If Me.OptionUBL_BE_3_0.Value = True Then
                If adminNoVat = True Then
                    peppolSuccess = XmlInvoiceGenerateUBLBE30NoVat((locPOSTVAKIN & "\" & idPdfForUbl))
                Else
                    peppolSuccess = XmlInvoiceGenerateUBLBE30((locPOSTVAKIN & "\" & idPdfForUbl))
                End If
            ElseIf Me.OptionPEPPOL_V3.Value = True Then
                peppolSuccess = XmlInvoiceGenerateUBLV3((locPOSTVAKIN & "\" & idPdfForUbl))
            Else
                MsgBox "logica", vbQuestion
            End If
            If peppolSuccess Then
            Else
                'MsgBox "Verkoopdocument blijft beschikbaar in " & vbCrLf & locPOSTVAKIN
            End If
            Kill locPOSTVAKIN & "\" & idPdfForUbl
        Else
            'customerPrefersEmail = False
        End If
    End If
End If

If useEmail = False Then
    mailAddressV224 = ""
    If String99(READING, 306) = "2" Then
        Msg = "Kopij van verkoopdocument staat klaar in de map Manueel voor afdruk."
        MsgBox Msg, vbInformation, "Preview uitgeschakeld in Setup"
    Else
        Mim.Report.Preview
    End If
Else
    If useEmail And customerPrefersEmail Then
        Screen.MousePointer = vbHourglass
        If Me.MPISessie.SessionID <> 0 Then
            On Local Error GoTo DVKMPIError
            'Compose new message
            Me.MPIBericht.Compose
            'Address message first only
            Me.MPIBericht.RecipDisplayName = rsMAR(TABLE_CUSTOMERS)("A100")
            Me.MPIBericht.RecipAddress = "SMTP:" & mailAddressV224
    
            'Resolve recipient name
            Me.MPIBericht.AddressResolveUI = True '=dialogbox, false = error genereren
            'MPIBericht.ResolveName
        
            'Create the message
            Me.MPIBericht.MsgSubject = pdfDOKUMENTTYPE + " " + dokumentSleutel
                
            Dim useDialogBox As Boolean
                
            useDialogBox = True
            If Dir(LOCATION_COMPANYDATA + "vpeSjbs\mailSjabloon-2.txt") = "" Then
                MsgBox "vpeSjbs\mailSjabloon-2.txt niet gevonden in de bedrijfsinhoudsopgave. Hierna wordt een voorbeeldsjabloon vanuit de programmainhoudsopgave gecopiëerd", vbInformation
                If Not CopyFile(App.path & "\VpeSjbs", LOCATION_COMPANYDATA + "vpeSjbs\", "mailSjabloon-2.txt") Then
                    MsgBox App.path & "\vpeSjbs\mailSjabloon-2.txt kan als voorbeelddocument niet vanuit de programmainhoudsopgave gecopiëerd worden.  Probeer eventueel manueel", vbInformation
                End If
            End If
            DoEvents
            emailTemplate = ""
            Ktrl = ScrLeesBestandAlleTekst(emailTemplate, LOCATION_COMPANYDATA + "vpeSjbs\mailSjabloon-2.txt")
            If Ktrl = 0 Then
                MsgBox "Onverwachte situatie", vbCritical
            End If
                
            emailTemplate = Replace(emailTemplate, "{docType}", pdfDOKUMENTTYPE)
            emailTemplate = Replace(emailTemplate, "{docNumber}", dokumentSleutel)
                
            sellerInfo = String99(READING, 52) + vbCrLf 'contact
            sellerInfo = sellerInfo + String99(READING, 46) + vbCrLf 'naam bedrijf
            sellerInfo = sellerInfo + String99(READING, 47) + vbCrLf 'straat en nr
            sellerInfo = sellerInfo + String99(READING, 48) + vbCrLf 'land pc en plaats
            sellerInfo = sellerInfo + "E. : " & String99(READING, 50) + vbCrLf 'mail
            sellerInfo = sellerInfo + "T. : " & String99(READING, 49) + vbCrLf 'telefoon
            sellerInfo = sellerInfo + String99(READING, 51)
                
            emailTemplate = Replace(emailTemplate, "{sellerInfo}", sellerInfo)
            Me.MPIBericht.MsgNoteText = emailTemplate
                
            'KtrlBox = MsgBox("Verstuurde E-mail aan " & vBibTekst(FlPartij, "#v224 #") & vbCr & vbCr & "Afdruk maken ?", vbQuestion + vbYesNo + vbDefaultButton2)
            Me.MPIBericht.AttachmentPathName = locPOSTVAKIN & "\" & idPdfForUbl
                
            'Send the message
            'On Error Resume Next
            Me.MPIBericht.Send useDialogBox
            If Err = 32001 Then
                Screen.MousePointer = vbNormal
                If Dir(locPOSTVAKIN & "\" & idPdfForUbl) = "" Then
                Else
                    Kill (locPOSTVAKIN & "\" & idPdfForUbl)
                End If
                Exit Function
            End If
            If Err Then
                'MsgBox Error
                If Dir(locPOSTVAKIN & "\" & idPdfForUbl) = "" Then
                Else
                    Kill (locPOSTVAKIN & "\" & idPdfForUbl)
                End If
                Mim.Report.Preview
            Else
                SnelHelpPrint "E-mail met succes verzonden.", BL_LOGGING
                'MsgBox "E-mail met succes.", vbInformation
                Mim.Report.CloseDoc
                If Dir(locPOSTVAKIN & "\" & idPdfForUbl) = "" Then
                Else
                    Kill (locPOSTVAKIN & "\" & idPdfForUbl)
                End If
            End If
        Else
            MsgBox "Uw E-mail systeem is niet geactiveerd in setup en de klant verkiest documenten via email.  Contacteer uw systeembeheerder voor bijkomende inlichtingen.", vbExclamation
            Mim.Report.Preview
        End If
        Screen.MousePointer = vbNormal
    End If
End If
Exit Function

DVKMPIError:
'MsgBox Error
Screen.MousePointer = vbNormal
MsgBox "Mailverzending afgebroken" & vbCrLf & vbCrLf & "Doe zelf verder het nodige", vbExclamation
If Dir(locPOSTVAKIN & "\" & idPdfForUbl) = "" Then
Else
    Kill (locPOSTVAKIN & "\" & idPdfForUbl)
End If
Mim.Report.Preview

End Function

Private Sub CmbExtraAfdruk_Click()

    'MsgBox "Extra afdruk is tijdelijk uitgeschakeld", vbInformation
    printAndSave = False
    reprintOnly = True
    Afdrukken
    
End Sub

Private Sub cmdLijst_Click()

Dim AAA As Variant
Dim KlantH As String
Dim DocH As String
Dim NogTeFaktureren As Boolean
Dim Aantal As Long
Dim recAantal As Long

If chkZonderRelatie.Visible = False Then
    NogTeFaktureren = False
ElseIf chkZonderRelatie.Value = "0" Then
    NogTeFaktureren = False
Else
    NogTeFaktureren = True
End If
   
Set rsDetail = New ADODB.Recordset
    
On Error Resume Next
Err = 0
rsDetail.CursorLocation = adUseClient
'Msg = "SELECT * FROM Allerlei WHERE v004 Like 'K%' AND v005 Like '" + dokumentType + "%' ORDER BY v004"
Msg = "SELECT * FROM Allerlei WHERE v005 Like '" + dokumentType + "%' ORDER BY v004, v005"
SnelHelpPrint Msg, BL_LOGGING
Screen.MousePointer = vbHourglass
rsDetail.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Screen.MousePointer = vbNormal
    Set rsDetail = Nothing
    Exit Sub
ElseIf rsDetail.RecordCount = 0 Then
    MsgBox "Geen dokumenten (meer) te vinden.", vbInformation
    Screen.MousePointer = vbNormal
    rsDetail.Close
    Set rsDetail = Nothing
    Exit Sub
ElseIf rsDetail.RecordCount > 1000 Then
    Msg = "Lijst maken voor " & rsDetail.RecordCount & " dokumenten." & vbCr
    Msg = Msg + "Bent U zeker" & vbCr
    KtrlBox = MsgBox(Msg, vbQuestion + vbDefaultButton2 + vbYesNo)
    If KtrlBox = vbNo Then
        rsDetail.Close
        Set rsDetail = Nothing
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
End If
SnelHelpPrint "Lijst samenstellen voor " & rsDetail.RecordCount & " lijnen", BL_LOGGING

If NogTeFaktureren = False Then
    Unload frmADOXlog
    frmADOXlog.Hide
    frmADOXlog.Caption = "documenten Lijst"
End If

If NogTeFaktureren Then
    lvDetail.ListItems.Clear
Else
    With frmADOXlog.lvDetail
        ' Add ColumnHeaders. The width of the columns is
        ' the width of the control divided by the number of
        ' ColumnHeader objects.
        .ColumnHeaders.Add , , "KlantInfo", 3120
        .ColumnHeaders.Add , , "document", 1245
        .ColumnHeaders.Add , , "Datum", 1065
        .ColumnHeaders.Add , , "Totaal Excl.", 1035, lvwColumnRight
        .ColumnHeaders.Add , , "Relatie tot dokumenten", 3120
        ' Set View property to Report.
        .View = lvwReport
    End With
End If
VSF_PRO = False
Dim itmX As ListItem
Do While Not rsDetail.EOF
    If NogTeFaktureren And Trim$(ADOBIB_TEXT(rsDetail("Memo"), "#v147 #")) <> "" Then
    Else
        If DocH <> rsDetail("v005") Then
            DocH = rsDetail("v005")
            If KlantH <> Mid(rsDetail("v004"), 2) Then
                KlantH = Mid(rsDetail("v004"), 2)
                If Not ADO_GET(TABLE_CUSTOMERS, 0, "=", Mid(rsDetail("v004"), 2)) Then
                    AAA = "- (" + Trim$(KlantH) + ")"
                Else
                    AAA = RV(rsMAR(TABLE_CUSTOMERS), "A100") + " (" + Trim$(KlantH) + ")"
                    ''Debug.Print AAA
                End If
            Else
                AAA = " "
            End If
            ' Add a ListItem object.
            If NogTeFaktureren Then
                Set itmX = lvDetail.ListItems.Add(, , AAA)
                itmX.SubItems(1) = ADOBIB_TEXT(rsDetail("Memo"), "#v033 #")
                itmX.SubItems(2) = DATE_TEXT(ADOBIB_TEXT(rsDetail("Memo"), "#v035 #"))
                itmX.SubItems(3) = ADOBIB_TEXT(rsDetail("Memo"), "#v137 #")
                itmX.SubItems(4) = ADOBIB_TEXT(rsDetail("Memo"), "#v147 #")
            Else
                Set itmX = frmADOXlog.lvDetail.ListItems.Add(, , AAA)
                itmX.SubItems(1) = ADOBIB_TEXT(rsDetail("Memo"), "#v033 #")
                itmX.SubItems(2) = DATE_TEXT(ADOBIB_TEXT(rsDetail("Memo"), "#v035 #"))
                itmX.SubItems(3) = ADOBIB_TEXT(rsDetail("Memo"), "#v137 #")
                itmX.SubItems(4) = ADOBIB_TEXT(rsDetail("Memo"), "#v147 #")
            End If
        End If
    End If
    rsDetail.MoveNext
Loop
rsDetail.Close
Set rsDetail = Nothing
If NogTeFaktureren Then
    SSTab1.Tab = 1
    Screen.MousePointer = vbNormal
    Me.VerkoopOptie.Item(0).Value = True
    Exit Sub
Else
    frmADOXlog.lblAantal(0).Caption = frmADOXlog.lvDetail.ListItems.Count
    Screen.MousePointer = vbNormal
    frmADOXlog.cmdOptie.Caption = "Verwijderen"
    frmADOXlog.Show 1
End If
Do While GridText <> ""
    Msg = dokumentType + Left(GridText, InStr(GridText, vbCr) - 1)
    GridText = Mid(GridText, InStr(GridText, vbCr) + 1)
    Msg = "DELETE * FROM Allerlei WHERE v005 = '" + Msg + "'"
    Err = 0
    On Error Resume Next
    adntDB.Execute Msg, recAantal
    If Err Then
        MsgBox Err
    Else
        MsgBox recAantal & " rijen werden verwijderd.", vbInformation
    End If
Loop
DirekteVerkoop.SetFocus


End Sub

Private Sub CmdOmschrijving_Click()

CmdOmschrijving.Default = True
If VerkoopDetail.Enabled Then
    VerkoopDetail_KeyPress Asc("O")
Else
    MsgBox "Eerst klant kiezen a.u.b !!!"
End If

End Sub

Private Sub cmdSQLInfo_Click()

If RV(rsKlant, "v253") <> "" Then
    KtrlBox = SQLPopUp(RV(rsKlant, "v253"), bstNaam(TABLE_CUSTOMERS), "A110", RV(rsKlant, "A110"))
Else
    cmdSQLInfo.Visible = False
End If

End Sub

Private Sub CmdStock_Click()

If VerkoopDetail.ListCount = 1 And adminNoVat = True Then
    MsgBox "Facturatie via omschrijving, bedrag geen BTW van toepassing", vbExclamation, "BTW Vrijstelling"
    Exit Sub
End If

    
CmdStock.Default = True
If VerkoopDetail.Enabled Then
    VerkoopDetail_KeyPress Asc("S")
Else
    MsgBox "Eerst klant kiezen a.u.b !!!"
End If

End Sub

Private Sub cmdSwitch_Click()
Dim TempoTel As Integer
Dim TempoVar As Variant
Dim MaskerEURBHmini As String

DIRECTSELL_STRING = cmdSwitch.Caption
MaskerEURBHmini = Mid(MASK_EURBH, 2)
If sMuntKlant = "BEF" Or sMuntKlant = "EUR" Then
'    If VerkoopDetail.ListCount Then
'        If cmdSwitch.Caption = "Ingave in BEF" Then
'            Msg = "Switch van BEF naar EUR.  Bij het later terugzetten naar BEF, bestaat de mogelijkheid van afrondingsverschillen indien meerdere lijnen in éénzelfde dokument aanwezig met oorspronkelijke cijfers in BEF na de komma.  De terugrekening vanuit de EURO naar BEF geeft in dat geval kans op afrondingsverschillen." & vbCr & vbCr & "Omrekening laten doorgaan ?"
'            KtrlBox = MsgBox(Msg, vbExclamation + vbYesNo + vbDefaultButton2)
'            If KtrlBox = vbNo Then Exit Sub
'        End If
'    End If
Else
    MsgBox "Switch niet mogelijk voor klanten buiten de E.U.  én enkel mogelijk indien klant met BEF of EUR code", vbInformation
    Exit Sub
End If
    
If cmdSwitch.Caption = "Ingave in EUR" Then
    'Stap 1: overschakeling cijfers van EUR naar BEF
    On Local Error Resume Next
    
    'For COUNT_TO = 1 To 3
        lblBTWBedrag1.Caption = Format(Round(CDbl(lblBTWBedrag1.Caption) * EURO, 0), "#,##0.00")
        lblBTWBedrag2.Caption = Format(Round(CDbl(lblBTWBedrag2.Caption) * EURO, 0), "#,##0.00")
        lblBTWBedrag3.Caption = Format(Round(CDbl(lblBTWBedrag3.Caption) * EURO, 0), "#,##0.00")
    'Next
    
    cmdSwitch.Caption = "Ingave in BEF"
    sMuntKlant = "BEF"
    For COUNT_TO = 0 To VerkoopDetail.ListCount - 1
       TempoVar = VerkoopDetail.List(COUNT_TO)
       If Right(TempoVar, 1) = "2" Then
       Else
           Mid(TempoVar, 42, 11) = Dec(Val(Mid(TempoVar, 42, 11)) * EURO, MaskerEURBHmini)
           Mid(TempoVar, 62, 12) = Dec(Val(Mid(TempoVar, 62, 12)) * EURO, MASK_EURBH)
           VerkoopDetail.List(COUNT_TO) = TempoVar
       End If
    Next
Else
    'overschakeling cijfers van BEF naar EUR
    'For COUNT_TO = 1 To 3
        lblBTWBedrag1.Caption = Format(Round(CDbl(lblBTWBedrag1.Caption) / EURO, 2), "#,##0.00")
        lblBTWBedrag2.Caption = Format(Round(CDbl(lblBTWBedrag2.Caption) / EURO, 2), "#,##0.00")
        lblBTWBedrag3.Caption = Format(Round(CDbl(lblBTWBedrag3.Caption) / EURO, 2), "#,##0.00")
    'Next
    cmdSwitch.Caption = "Ingave in EUR"
    sMuntKlant = "EUR"
    For COUNT_TO = 0 To VerkoopDetail.ListCount - 1
        TempoVar = VerkoopDetail.List(COUNT_TO)
        If Right(TempoVar, 1) = "2" Then
        Else
            Mid(TempoVar, 42, 11) = Dec(Val(Mid(TempoVar, 42, 11)) / EURO, MaskerEURBHmini)
            Mid(TempoVar, 62, 12) = Dec(Val(Mid(TempoVar, 62, 12)) / EURO, MASK_EURBH)
            VerkoopDetail.List(COUNT_TO) = TempoVar
        End If
    Next
End If
TekstInfo5.Text = Dec$(dMuntK, "##0.########")
DIRECTSELL_STRING = cmdSwitch.Caption
RefreshBTW

End Sub

Private Sub CmdTekst_Click()

If VerkoopDetail.ListCount = 1 And adminNoVat = True Then
    MsgBox "Facturatie via omschrijving, bedrag geen BTW van toepassing", vbExclamation, "BTW Vrijstelling"
    Exit Sub
End If

CmdTekst.Default = True
If VerkoopDetail.Enabled Then
    VerkoopDetail_KeyPress Asc("T")
Else
    MsgBox "Eerst klant kiezen a.u.b !!!"
End If

End Sub

Private Sub CreditNota_Click()

Dim peppolCtrl As Boolean

If KlantInfo.Caption <> "" Then
    If customerCompanyId = "" Then
    Else
        peppolCtrl = checkForB2BInvoice()
    End If
End If

If CreditNota.Value = 0 Then
Else
    VerkoopOptie(0).Value = 1
    VerkoopOptie(0).Enabled = False
    VerkoopOptie(1).Enabled = False
    VerkoopOptie(2).Enabled = False
    dokumentSleutel = SleutelDok(13)
    TekstInfo3 = dokumentSleutel
    Vr = 13
    CreditNota.Enabled = False
End If
DirekteVerkoop.Caption = vSet("Verkoopverrichting", 28) + "(" + dokumentSleutel + ")"

End Sub

Private Sub pdfDrukAf()

Dim tabV As Integer
Dim hTAB    As Integer
ReDim rSip(6) As Variant
ReDim sSip(6) As Variant
ReDim VeldInfo(9) As String
Dim A As String

ReDim dVeldInfo(7) As Double
Dim TotaalBedrag As Double
Dim BtwTekst As String * 4

Dim OMSTab As Integer
Dim tPagina As Integer
Dim Pagina As Integer
Dim tSip As Integer
Dim iRNTxt As Integer
Dim Teltxt As Integer
Dim strMeerLijn As String

Dim BedragTxt As String
Dim rNTTxt As String
Dim rNTTxt2 As String
Dim ReferteTxt As String

Dim T                   As Integer
Dim TT                  As Integer
Dim Taal                As String
Dim FlTemp              As Integer
ReDim rft(10) As String
Dim rnr                 As String * 13
Dim sy                  As String
Dim sy2                 As String

Dim XVan                As Single
Dim XTot                As Single
Dim YVan                As Single
Dim YTot                As Single
Dim X                   As Integer
Dim MeerLijn            As Integer

Dim Nog As Integer
Dim NogString As String

Dim ktrlHier As Double
Dim adresString As String
Dim ovsDefinitie As String
Dim aantalPaginas As Integer


'On Local Error GoTo FileHandler
If Overschrijvingsstrook.Value Then
    pdfOVSStrook = 8.2
Else
    pdfOVSStrook = 0
End If

MeerLijn = Val(String99(READING, 72))
Taal = RV(rsKlant, "A10C")
FlTemp = FreeFile

Open PROGRAM_LOCATION + "Def\f0" + Taal + ".srn" For Input As FlTemp
    T = 0
    While Not EOF(FlTemp)
        Input #FlTemp, rft(T)
        T = T + 1
    Wend
Close FlTemp

TRYIBANAGAIN:
ovsDefinitie = "FORM-IBAN.TXT"
If Dir(LOCATION_COMPANYDATA + ovsDefinitie) = "" Then
    Msg = LOCATION_COMPANYDATA & ovsDefinitie & " niet gevonden" & vbCrLf
    Msg = Msg & "Hierna volgt een voorbeeldsjabloon vanuit de programmainhoudsopgave" & vbCrLf & vbCrLf
    Msg = Msg & "Wijzig met uw gegevens via kladblok. Kladblok wordt hierna geladen."
    MsgBox Msg, vbInformation
    If Not CopyFile(App.path & "\Def", LOCATION_COMPANYDATA, "FORM-IBAN.TXT") Then
        MsgBox App.path & "\vpeSjbs\mailSjabloon-2.txt kan als voorbeelddocument niet vanuit de programmainhoudsopgave gecopiëerd worden.  Probeer eventueel manueel", vbInformation
    Else
        On Local Error Resume Next
        X = Shell("notepad.exe " + LOCATION_COMPANYDATA + ovsDefinitie, 1)
    End If
    GoTo TRYIBANAGAIN
Else
    FlTemp = FreeFile
    ovsDefinitie = "FORM-IBAN.TXT"
    Open LOCATION_COMPANYDATA + ovsDefinitie For Input As FlTemp
        Line Input #FlTemp, sSip(0)
        Line Input #FlTemp, sSip(1)
        Line Input #FlTemp, sSip(2)
        Line Input #FlTemp, sSip(3)
        Line Input #FlTemp, sSip(4)
        Line Input #FlTemp, sSip(5)
        Line Input #FlTemp, sSip(6)
    Close FlTemp
End If

'On Local Error GoTo PrintHandler
aantalPaginas = 1 'CInt((Val(RV(rsKlant, "vs07")) + Val(Right(AantalEx, 1))) / Val(Left(AantalEx, 1)))
        
For tPagina = 1 To aantalPaginas
    Pagina = 0
    GoSub pdfKopBalk
    GoSub pdfDetailLijnen
    GoSub pdfVoetTekst
    GoSub qrCodeGen
    If Overschrijvingsstrook.Value Then
        GoSub pdfOverschrijvingsstrook
    End If
    If tPagina = aantalPaginas Then
    Else
        Ktrl = Mim.Report.PageBreak
    End If
Next tPagina
If Me.cbAV.Value = vbChecked Then
    'Ktrl = Mim.Report.PageBreak
    Ktrl = GenerateVpeDoc(LOCATION_COMPANYDATA & "\vpeSjbs\", "VpeAV.txt")
End If
Exit Sub

pdfKopBalk:
Pagina = Pagina + 1
If Vr = 11 Then
    pdfDOKUMENTTYPE = Left(rft(0), 14)
ElseIf Vr = 13 Then
    pdfDOKUMENTTYPE = Right(rft(0), 14)
ElseIf Vr = 73 Then
    pdfDOKUMENTTYPE = "Bestelbon/Leveringsbon"
Else
    pdfDOKUMENTTYPE = "Offerte"
End If

If Val(RV(rsKlant, "A102")) = 0 Then
    rSip(0) = RV(rsKlant, "A100") + " " + RV(rsKlant, "A101")
Else
    rSip(0) = Mid(fmarBoxText("003", Taal, RV(rsKlant, "A102")), 4, 10) + " " + RV(rsKlant, "A100") + " " + RV(rsKlant, "A101")
End If

customerName = Trim(rSip(0))
If InStr(customerName, "&") Then
    customerName = Replace(customerName, "&", "&amp;")
End If
customerRegistrationName = customerName

If KontaktPersoon = 1 Then
    If Val(RV(rsKlant, "vs01")) = 0 Then
        rSip(1) = RV(rsKlant, "A125") + " " + RV(rsKlant, "A127")
    Else
        rSip(1) = Mid(fmarBoxText("003", Taal, RV(rsKlant, "vs01")), 4, 10) + " " + RV(rsKlant, "A125") + " " + RV(rsKlant, "A127")
    End If
Else
    rSip(1) = ""
End If

rSip(2) = RV(rsKlant, "A104") & " " & RV(rsKlant, "A105") & " " & RV(rsKlant, "A106")
customerStreetName = Trim(rSip(2))
If InStr(customerStreetName, "&") Then
    customerStreetName = Replace(customerStreetName, "&", "&amp;")
End If

customerCityName = Trim(RV(rsKlant, "A108"))
customerPostalZone = Trim(RV(rsKlant, "A107"))
customerCountryCode = Trim(RV(rsKlant, "v150"))
rSip(4) = RV(rsKlant, "A109") & " " & RV(rsKlant, "A107") & " " & RV(rsKlant, "A108")

With Mim.Report
    .FontName = "Courier New"
    .FontSize = 10
    .TextBold = True
    .TextColor = 0 'zwart
    .nTopMargin = 1
    .nLeftMargin = 0.5
    .nRightMargin = 0.5
    .PenSize = 0.01
End With
GoSub pdfPrintKopTekst

With Mim.Report
    .FontName = "Courier New"
    .FontSize = 10
    .TextBold = True
End With

Mim.Report.PenSize = 0.01
pdfY = Mim.Report.PrintBox(1.5, pdfVsoftVanaf, UCase(pdfDOKUMENTTYPE))

invoiceCustomerNumber = Trim(RV(rsKlant, "A110"))
invoiceNumber = Trim(dokumentSleutel) + "-" + invoiceCustomerNumber
paymentTerms = invoiceNumber
orderReference = invoiceNumber
paymentID = invoiceNumber

pdfY = Mim.Report.VPEPRINT(8.5, pdfVsoftVanaf, dokumentSleutel & " / " & Str(Pagina))
pdfY = Mim.Report.VPEPRINT(1.5, pdfY, vbCrLf)

If Trim(RV(rsKlant, "A161")) = "" Then
    customerCompanyIdExtended = ""
Else
    customerCompanyId = Trim(RV(rsKlant, "v404"))
    customerCompanyIdExtended = Trim(RV(rsKlant, "v150")) + customerCompanyId
    customerVatNumber = Trim(RV(rsKlant, "A161"))
    If customerVatNumber <> "" Then customerTaxScheme = "VAT"
End If

pdfY = Mim.Report.PrintBox(1.5, pdfY, rft(1) & vbCrLf & vSet(RV(rsKlant, "A110"), 12) & " " & vSet(customerCompanyIdExtended, 14) & " " & TekstInfo0.Text & " " & TekstInfo1.Text)
pdfY = Mim.Report.VPEPRINT(1.5, pdfY, vbCrLf)

If MeerLijn = 1 Then
    Mim.Report.FontSize = 7.2
Else
    Mim.Report.FontSize = 10
End If

If Left(String99(READING, 74), 1) = "2" Then
    If MeerLijn = 1 Then
        strMeerLijn = Space$(Len(rft(2))) + " "
        'OMSTab = Len(rft(2)) + 1
    End If
ElseIf Left(String99(READING, 75), 1) = "2" Then
    strMeerLijn = Left(rft(2), 13) + Space$(13) + " "
    If MeerLijn = 1 Then
        'OMSTab = 14
    Else
        strMeerLijn = strMeerLijn + vbCrLf
        'OMSTab = 1
    End If
Else
    strMeerLijn = rft(2) + " "
    If MeerLijn = 1 Then
    Else
        strMeerLijn = strMeerLijn + vbCrLf
    End If
End If
If Left(String99(READING, 76), 1) = "2" Then
    Mid(rft(10), 63, 2) = "  "
End If
strMeerLijn = strMeerLijn + rft(10)
pdfY = Mim.Report.PrintBox(1.5, pdfY, strMeerLijn)
pdfY = Mim.Report.VPEPRINT(1.5, pdfY, vbCrLf)
Return

pdfDetailLijnen:
Mim.Report.TextBold = True
For TT = 0 To VerkoopDetail.ListCount - 1
    GridText = VerkoopDetail.List(TT)
    
    If Right(GridText, 1) = "2" Then
        pdfY = Mim.Report.VPEPRINT(1.7, pdfY, Left(GridText, 75))
        'pdfY = Mim.Report.VpePrint(1.7, pdfY, vbCrLf)
        GoTo pdfKontroleLijn
    Else
        GoSub pdfFilterVelden
    End If
    If Right(GridText, 1) = "0" Then
        If Left(String99(READING, 74), 1) = "2" Then
            If MeerLijn = 1 Then
                strMeerLijn = Space$(Len(rft(2))) + " "
            End If
        ElseIf Left(String99(READING, 75), 1) = "2" Then
            strMeerLijn = vSet(VeldInfo(0), 13)
            If MeerLijn = 1 Then
                strMeerLijn = strMeerLijn + Space$(13) + " "
                pdfY = Mim.Report.VPEPRINT(1.7, pdfY, strMeerLijn)
            Else
                pdfY = Mim.Report.VPEPRINT(1.7, pdfY, strMeerLijn)
                pdfY = Mim.Report.VPEPRINT(1.7, pdfY, vbCrLf)
            End If
        Else
            If MeerLijn = 1 Then
            Else
                pdfY = Mim.Report.VPEPRINT(1.7, pdfY, vSet(VeldInfo(0), 13) + " " + Dec$((dVeldInfo(2)), "###0.0") + " " + vSet(Mid(fmarBoxText("004", Taal, VeldInfo(8)), 4, 5), 5) + " ")
            End If
        End If
    End If
    If MeerLijn = 1 Then
        strMeerLijn = vSet(VeldInfo(0), 13) + " " + Dec$((dVeldInfo(2)), "###0.0") + " " + vSet(Mid(fmarBoxText("004", Taal, VeldInfo(8)), 4, 5), 5) + " "
    Else
        strMeerLijn = ""
        'pdfY = Mim.Report.VpePrint(1.5, pdfY, "")
    End If
    strMeerLijn = strMeerLijn + vSet(VeldInfo(1), 40) + " " + Dec$(dVeldInfo(4) / dMuntK, "######0.000") + " " + Dec$((dVeldInfo(6)), vkMaskAantal) + " "
    If Left(String99(READING, 76), 1) = "2" Then
        strMeerLijn = strMeerLijn + "    "
    Else
        strMeerLijn = strMeerLijn + Dec$((dVeldInfo(5)), "##0") + " "
    End If
    
    BtwTekst = ""
    Mid(BtwTekst, 1) = Mid(fmarBoxText("002", "2", VeldInfo(9)), 4, 4)
    
    If InStr(cmdSwitch.Caption, "EUR") Then
        strMeerLijn = strMeerLijn + BtwTekst + " " + Dec$(dVeldInfo(7) / dMuntK, "########0.00") + vbCrLf
    ElseIf dMuntK <> 1 Then
        strMeerLijn = strMeerLijn + BtwTekst + " " + Dec$(dVeldInfo(7) / dMuntK, "########0.00") + vbCrLf
    Else
        strMeerLijn = strMeerLijn + BtwTekst + " " + Dec$(Val(Format(dVeldInfo(7) / dMuntK, "#")), "########0.00") + vbCrLf
    End If
    pdfY = Mim.Report.VPEPRINT(1.7, pdfY, strMeerLijn)


pdfKontroleLijn:
If pdfY >= pdfVsoftTot - 3.2 - pdfOVSStrook Then
    GoSub pdfOnderKant
    GoSub pdfVoetTekst
    Ktrl = Mim.Report.PageBreak
    'MsgBox "stop voor nieuwe pdf pagina nog te verbeteren"
    GoSub pdfKopBalk
End If
Next

pdfOnderKant:
pdfY = Mim.Report.VPEPRINT(1.5, pdfVsoftTot - 2.3 - pdfOVSStrook, vbCrLf)

Mim.Report.FontSize = 10
sy = "####0.00"
sy2 = MASK_EUR

Mid(rft(5), 25, 4) = Mid(fmarBoxText("002", "2", "1"), 4, 4)
Mid(rft(5), 35, 4) = Mid(fmarBoxText("002", "2", "2"), 4, 4)
Mid(rft(5), 45, 4) = Mid(fmarBoxText("002", "2", "3"), 4, 4)
Mim.Report.TextBold = False

    strMeerLijn = rft(5) + "  " + rft(6) + Dec$(CDbl(LblEx2Btw) / dMuntK, sy2) + vbCrLf
    strMeerLijn = strMeerLijn + " " + Dec(0, sy) + "  " + Dec(0, sy) + "  " + Dec(BTWEuroBedrag(1) / dMuntK, sy)
    strMeerLijn = strMeerLijn + "  " + Dec(BTWEuroBedrag(2) / dMuntK, sy) + "  " + Dec(BTWEuroBedrag(3) / dMuntK, sy) + "  " + rft(7)
    strMeerLijn = strMeerLijn + Dec((TotaalBTW) / dMuntK, sy2) + vbCrLf

    strMeerLijn = strMeerLijn + " " + Dec(TotaalUitvoer / dMuntK, sy) + "  " + Dec(BTWEuroBasis(0) / dMuntK, sy) + "  " + Dec(BTWEuroBasis(1) / dMuntK, sy)
    strMeerLijn = strMeerLijn + "  " + Dec(BTWEuroBasis(2) / dMuntK, sy) + "  " + Dec(BTWEuroBasis(3) / dMuntK, sy) + "  " + rft(8) + sMuntKlant + ":"
    TotaalBedrag = TotaalBTW + CDbl(LblEx2Btw)
    strMeerLijn = strMeerLijn + Dec(TotaalBedrag / dMuntK, sy2)
    
    taxGlobalTotalAmount = Trim(Dec((TotaalBTW), sy2))
    taxableGlobalTotalAmount = Trim(Dec(CDbl(LblEx2Btw.Caption), sy2))
    
    
If Medekontraktant.Value Then
    strMeerLijn = strMeerLijn + vbCrLf + "  " + rft(4)
End If
pdfY = Mim.Report.PrintBox(3, pdfY, strMeerLijn)

Return

pdfPrintKopTekst:
If Vr = 11 Then
    'Faktuur
    pdfPrintUserDef "1" + Taal + "0", pdfOVSStrook
ElseIf Vr = 13 Then
    'Creditnota
    pdfPrintUserDef "1" + Taal + "0", pdfOVSStrook
ElseIf Vr = 73 Then
    'Bestelbon/Leveringsbon
    pdfPrintUserDef "1" + Taal + "1", pdfOVSStrook
Else
    'Offerte
    pdfPrintUserDef "1" + Taal + "3", pdfOVSStrook
End If

If usrLicentieInfo <> "" Then
    pdfY = Mim.Report.VPEPRINT(0.6, 0.6, usrLicentieInfo)
End If

With Mim.Report
    .FontSize = 10
    .TextBold = True
    .TextItalic = False
    .TextUnderline = False
    .TextAlignment = 0
    .TextColor = 0 'zwart
End With

adresString = ""
For tSip = 0 To 4
    adresString = adresString & UCase(rSip(tSip)) & vbCrLf
    'Printer.Print UCase$(rSip(tSip)); vbCrLf;
    'Printer.CurrentX = BeginXbox + 350
Next
ktrlHier = Mim.Report.Write(pdfadresXpos, pdfadresYpos, pdfadresXpos2, pdfadresYpos2, adresString)
Return

pdfVoetTekst:
If OVSStrooklijnen Then
    GoSub pdfOverschrijvingsstrook
End If
Return

pdfFilterVelden:
VeldInfo(1) = Left(GridText, 40)
GridText = Right(GridText, Len(GridText) - 41)
dVeldInfo(4) = Val(Left(GridText, 11))
GridText = Right(GridText, Len(GridText) - 12)
dVeldInfo(6) = Val(Left(GridText, 7))
GridText = Right(GridText, Len(GridText) - 8)
dVeldInfo(7) = Val(Left(GridText, 12))
GridText = Right(GridText, Len(GridText) - 13)
dVeldInfo(2) = Val(Left(GridText, 6))
GridText = Right(GridText, Len(GridText) - 7)
VeldInfo(8) = Left(GridText, 1)
GridText = Right(GridText, Len(GridText) - 2)
dVeldInfo(5) = Val(Left(GridText, 3))
GridText = Right(GridText, Len(GridText) - 4)
VeldInfo(9) = Left(GridText, 1)
GridText = Right(GridText, Len(GridText) - 2)
'If VeldInfo(9) = "6" Then
'    VeldInfo(9) = "0"
'End If
dVeldInfo(3) = Val(Left(GridText, 7))
GridText = Right(GridText, Len(GridText) - 8)
VeldInfo(0) = Left(GridText, 13)
Return

qrCodeGen:
If isDOM Then
Else
    If Me.cbQRCode.Value = vbUnchecked Then
    Else
        Dim qrFileName As String
        Dim findQrFile As String
    
        Dim betaalinfo As String
    
        Dim serviceTagValue As String
        Dim versionValue As String
        Dim charactersetValue As String
        Dim identificationValue As String
        Dim bicValue As String
        Dim nameValue As String
        Dim ibanValue As String
        Dim amountValue As String
        Dim purposeValue As String
        Dim referenceValue As String
        Dim remittanceValue As String
        Dim informationValue As String
    
    
        serviceTagValue = "BCD" + vbCrLf
        versionValue = "001" + vbCrLf
        charactersetValue = "1" + vbCrLf
        identificationValue = "SCT" + vbCrLf
        bicValue = sSip(1) + vbCrLf
        nameValue = sSip(2) + vbCrLf
        ibanValue = sSip(0) + vbCrLf
        amountValue = "EUR" + Trim(Dec((TotaalBedrag), "#######0.00")) + vbCrLf
        purposeValue = "GDDS" + vbCrLf
        referenceValue = dokumentSleutel + vbCrLf
        remittanceValue = vbCrLf
        informationValue = vbCrLf
    
        DoEvents
        FileNameQR = dokumentSleutel
        qrFileName = LOCATION_COMPANYDATA + "BMP-qr\" + FileNameQR + ".bmp"
        
        ShowAndSaveQR.Hide
        DoEvents
        ShowAndSaveQR.Text1.Text = serviceTagValue + versionValue + charactersetValue + identificationValue + bicValue + nameValue + ibanValue + amountValue + purposeValue + referenceValue + remittanceValue + informationValue
        DoEvents
        Unload ShowAndSaveQR
        DoEvents
        
        Mim.Report.PictureBestFit = False
        findQrFile = Dir(qrFileName)
        DoEvents
        If findQrFile = "" Then
            MsgBox "error cannot find " + qrFileName
        Else
            'onder addres
            'Ktrl = Mim.Report.Picture(17, 8, 19, 10, qrFileName)
        
            If Overschrijvingsstrook.Value Then
                'met ovstrook
                Ktrl = Mim.Report.Picture(16.8, 18, 18.8, 20, qrFileName)
            Else
                'zonder ovstrook
                Ktrl = Mim.Report.Picture(17, 26.2, 19, 28.2, qrFileName)
            End If
            DoEvents
        End If
    End If
End If

Return

pdfOverschrijvingsstrook:
On Local Error GoTo 0
Mim.Report.FontName = "Courier New"
Mim.Report.FontSize = 12
Mim.Report.TextBold = True

If isDOM Then
    rNTTxt = "********.**"
Else
    rNTTxt = Dec((TotaalBedrag), "#######0.00")
End If
GoSub pdfSpatieren
Mid(rNTTxt2, 17, 1) = " "
Dim tmppdfY As Double

pdfY = Mim.Report.VPEPRINT(15, 22, rNTTxt2)     'bedrag
Mim.Report.TextBold = False
    
    rNTTxt = Left(UCase(rSip(0)), 26) 'Klant naam1
    GoSub pdfSpatieren
    pdfY = Mim.Report.VPEPRINT(3.6, 23.7, rNTTxt2 & vbCrLf)
    
    rNTTxt = Left(UCase(rSip(2)), 26) 'Klant straat
    GoSub pdfSpatieren
    pdfY = Mim.Report.VPEPRINT(3.6, pdfY, rNTTxt2 & vbCrLf)
    
    rNTTxt = Left(UCase(rSip(4)), 26) 'Klant plaats
    GoSub pdfSpatieren
    pdfY = Mim.Report.VPEPRINT(3.6, pdfY, rNTTxt2 & vbCrLf)
        
    Mim.Report.TextBold = True
    If isDOM = True Then
        rNTTxt = "*domiciliëring: enkel ter info*"
    Else
        rNTTxt = sSip(0)                                'IBANbedrijf
    End If
    GoSub pdfSpatieren
    pdfY = Mim.Report.VPEPRINT(3.6, 25.3, rNTTxt2 & vbCrLf & vbCrLf)
    
    rNTTxt = sSip(1)                                'BICbedrijf
    GoSub pdfSpatieren
    pdfY = Mim.Report.VPEPRINT(3.6, pdfY, rNTTxt2 & vbCrLf & vbCrLf)
    
    Mim.Report.TextBold = False
    For T = 2 To 4
        rNTTxt = sSip(T) 'ADRESbedrijf
        GoSub pdfSpatieren
        pdfY = Mim.Report.VPEPRINT(3.6, pdfY, UCase(rNTTxt2))
    Next
        
    rNTTxt = dokumentSleutel
    GoSub pdfSpatieren
    Mim.Report.TextBold = True
    pdfY = Mim.Report.VPEPRINT(3.6, pdfY, rNTTxt2)
Return

pdfSpatieren:
iRNTxt = Len(rNTTxt)
rNTTxt2 = ""
For Teltxt = 1 To iRNTxt
    rNTTxt2 = rNTTxt2 + Mid(rNTTxt, Teltxt, 1) + " "
Next
Return

pdfPrintHandler:
MsgBox "Stopkode extern toestel." + vbCrLf + vbCrLf + "Kontroleer, of rapporteer aan Vsoft.", 0, "Onverwachte situatie"
Resume

pdfFileHandler:
MsgBox "Onverwachte fout bij openen bestand." + vbCrLf + vbCrLf + "Stopkode : " + Format(Err) + vbCrLf + vbCrLf + "Rapporteer aan Vsoft of kontroleer uw setup."
Ktrl = 100
Exit Sub
Resume

End Sub


' Make sure to add a reference to "Microsoft XML, v6.0" in your VB6 project
Private Function SaveBinaryDataToXML(ByVal sInputFile As String, ByVal sOutputXML As String) As Boolean
    Dim bData() As Byte
    Dim nFileNum As Integer
    
    SaveBinaryDataToXML = False
    ' Step 1: Read the binary data from the input file
    nFileNum = FreeFile
    Open sInputFile For Binary As #nFileNum
        ReDim bData(LOF(nFileNum) - 1)
        Get #nFileNum, , bData
    Close #nFileNum

    ' Step 2: Convert the byte array to a Base64 string using MSXML
    Dim xmlDoc As MSXML2.DOMDocument60
    Set xmlDoc = New MSXML2.DOMDocument60
    
    Dim base64Elem As IXMLDOMElement
    Set base64Elem = xmlDoc.createElement("BinaryData")
    
    ' Assign the binary data to the element.
    ' MSXML will convert it to Base64 since we specify its DataType property.
    base64Elem.dataType = "bin.base64"
    base64Elem.nodeTypedValue = bData
    
    Dim base64Str As String
    base64Str = base64Elem.Text   ' This is the Base64 encoded string

    ' Step 3: Build the XML document and insert the Base64 string
    Dim rootElem As IXMLDOMElement
    Set rootElem = xmlDoc.createElement("Root")
    xmlDoc.appendChild rootElem

    Dim dataElem As IXMLDOMElement
    Set dataElem = xmlDoc.createElement("FileData")
    dataElem.Text = base64Str   ' Insert the Base64 string as the text of the element
    rootElem.appendChild dataElem

    ' Save the XML to a file
    xmlDoc.Save sOutputXML
    
    MsgBox "Binary data has been encoded to Base64 and saved in XML format.", vbInformation, "Operation Complete"
    SaveBinaryDataToXML = True
    
End Function

Private Function GetBase64Str(ByVal sInputFile As String) As String
    Dim bData() As Byte
    Dim nFileNum As Integer
    
    GetBase64Str = ""
    ' Step 1: Read the binary data from the input file
    nFileNum = FreeFile
    Open sInputFile For Binary As #nFileNum
        ReDim bData(LOF(nFileNum) - 1)
        Get #nFileNum, , bData
    Close #nFileNum

    ' Step 2: Convert the byte array to a Base64 string using MSXML
    Dim xmlDoc As MSXML2.DOMDocument60
    Set xmlDoc = New MSXML2.DOMDocument60
    
    Dim base64Elem As IXMLDOMElement
    Set base64Elem = xmlDoc.createElement("BinaryData")
    
    ' Assign the binary data to the element.
    ' MSXML will convert it to Base64 since we specify its DataType property.
    base64Elem.dataType = "bin.base64"
    base64Elem.nodeTypedValue = bData
    
    Dim base64HereStr As String
    
        
    base64HereStr = base64Elem.Text   ' This is the Base64 encoded string
    GetBase64Str = base64HereStr
    
End Function


Private Sub Form_Activate()

DoEvents
DirekteVerkoop.SetFocus

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

If (Right(LOCATION_COMPANYDATA, 5) = "\098\" Or Right(LOCATION_COMPANYDATA, 5) = "\099\") Then
    Me.TextBoxWarningTestCompany.Visible = True
End If


ret = SHGetFolderPath(0, 0, 0, 0, path)
desktopLocatie = Left(path, InStr(path, Chr(0)) - 1)

If Trim(String99(READING, 20)) = "7" Then
    adminNoVat = True
Else
    adminNoVat = False
End If

If LaadTekst("dnnInstellingen", "Cloud") = "" Then
    MsgBox "Nieuwe PC of nog geen instellingen voor beheer documentenstroom. Geef minstens een standaardlocatie voor postvak in a.u.b.", vbInformation
Else
    locPOSTVAKIN = LaadTekst("dnnInstellingen", "Mario")
End If

Dim TestString As String
invoiceTypeCode = "380" '(see https://docs.peppol.eu/poac/pint/pint/trn-invoice/codelist/UNCL1001-inv/)
TestString = ZoekEnPlaats(Me.cbUncl1001, "NTKB5001", 0, 0, invoiceTypeCode)
    
With lvDetail
    ' Add ColumnHeaders. The width of the columns is
    ' the width of the control divided by the number of
    ' ColumnHeader objects.
    .ColumnHeaders.Add , , "KlantInfo", 3120
    .ColumnHeaders.Add , , "document", 1245
    .ColumnHeaders.Add , , "Datum", 1065
    .ColumnHeaders.Add , , "Totaal Excl.", 1035, lvwColumnRight
    .ColumnHeaders.Add , , "Relatie tot documenten", 3120
    ' Set View property to Report.
    .View = lvwReport
End With

'uitsluitend EURO vanaf 2011
cmdSwitch.Caption = "Ingave in EUR"
If bhEuro Then
    cmdSwitch.Caption = "Ingave in EUR"
Else
    bhEuro = True
End If

Dim Tel As Integer

Top = 0
Left = 0
Screen.MousePointer = vbHourglass

dMuntK = 1
sMunt = "BEF"

For Tel = 16 To 19
    rbtwVAK(Tel - 16) = String99(READING, Tel)
    rbtwVAK(Tel - 12) = String99(READING, Tel + 6)
Next

MaskAantal:
vkMaskAantal = String99(READING, 300)
If Len(Trim$(vkMaskAantal)) <> 7 Then
    SS99 "0000.00", 300
    MsgBox "Masker voor aantal is ingesteld als '0000.00'  Wijzig via setup indien nodig.", vbInformation
    GoTo MaskAantal
End If
Mid(vkMaskAantal, 1, InStr(vkMaskAantal, ".") - 2) = String(InStr(vkMaskAantal, ".") - 2, "#")

AantalEx = String99(READING, 185)

KontaktPersoon = Val(String99(READING, 201))
If Val(String99(READING, 202)) = 1 Then
    Overschrijvingsstrook.Value = 1
Else
    Overschrijvingsstrook.Value = 0
End If
If Val(String99(READING, 203)) = 1 Then
    SteedsDrukken = True
Else
    SteedsDrukken = False
End If
Vr = 11
Dim X As Integer
Schoon

supplierTaxScheme = ""
supplierRegistrationId = Trim(String99(READING, 292))
supplierVatNumber = Trim(String99(READING, 51))
If supplierVatNumber = supplierRegistrationId Then supplierTaxScheme = "VAT"

supplierRegistrationName = Trim(String99(READING, 46))
If InStr(supplierRegistrationName, "&") Then
    supplierRegistrationName = Replace(supplierRegistrationName, "&", "&amp;")
End If

If Left(String99(READING, 20), 1) = "4" Then
    ForFait = 1
    MsgBox "Verkoopfakturatie voor forfaitair BTW SYSTEEM."
    TekstInfo3.Visible = True
Else
    ForFait = 0
End If
KlantRekening = String99(READING, 9)
TekstInfo2.Text = KlantRekening

If String99(READING, 290) = "1" Then
    Me.cbQRCode.Value = vbChecked
Else
    Me.cbQRCode.Value = vbUnchecked
End If

If String99(READING, 299) = "2" Then
    SS99 "1", 299
End If

Dim TempoKLS As String
Dim TempoDOK As String

If InStr(GridText9, vbTab) <> 0 Then
    TempoKLS = Mid(GridText9, 1, InStr(GridText9, vbTab) - 1)
    GridText9 = Mid(GridText9, InStr(GridText9, vbTab) + 1)
    XLogKey = TempoKLS
    InstalKlant
    Select Case Left(GridText9, 2)
        Case "13", "14", "15"
            dokumentType = Left(GridText9, 2)
            TempoDOK = Mid(GridText9, InStr(GridText9, vbTab) + 1)
            If Len(TempoDOK) <> 8 Then
                VerkoopOptie_Click 0
            Else
                XLogKey = TempoDOK
                LaadHetdokument
            End If
    End Select
End If
Screen.MousePointer = vbNormal

If String99(READING, 21) = "1" Then
'    MsgBox "Outlook Email wordt niet meer gebruikt in marIntegraal versie 10.4 of hoger.  Schakel deze instelling uit en gebruik VPE/PDF en/of de mario mailspooler eens alle toestellen beschikken over marIntegraal 10.4 of hoger", vbInformation
    useEmail = True
    On Local Error Resume Next
    Me.MPISessie.SignOff
    On Local Error Resume Next
    Me.MPISessie.SignOn
    If Err Then
        MsgBox Err.Description
    Else
        On Local Error GoTo 0
        Me.MPIBericht.SessionID = Me.MPISessie.SessionID
        SnelHelpPrint "E-mail sessie met succes opgestart. IDkode :" + Format(Me.MPISessie.SessionID), BL_LOGGING
    End If
End If
    
On Error Resume Next
DirekteVerkoop.SetFocus

If checkSetUp = False Then
    Me.KlantAktiveren.Enabled = False
End If

End Sub

Function checkSetUp() As Boolean

    Dim numberFalse As Integer
    
    checkSetUp = False
    If supplierVatNumber = "" Then
        MsgBox "controleer uw BTW taxschema a.u.b.", vbExclamation
        numberFalse = numberFalse + 1
    End If
    If Mid(String99(READING, 48), 5, 1) = " " Then
        supplierPostalZone = Mid(String99(READING, 48), 1, 4)
        supplierCityName = Trim(Mid(String99(READING, 48), 5))
    Else
        MsgBox "Controleer setup en parameters a.u.b." & vbCrLf & vbCrLf & "O.a. postcode uit 4 cijfers gevolgd door een spatie.", vbExclamation
        numberFalse = numberFalse + 1
    End If
    supplierStreetName = Trim(String99(READING, 47))
    If InStr(supplierStreetName, "&") Then
        supplierStreetName = Replace(supplierStreetName, "&", "&amp;")
    End If
    
    supplierCountryCode = "BE"
    supplierCompanyId = supplierRegistrationId
    supplierCompanyIdExtended = "BE" + supplierCompanyId
    
    supplierTelephone = String99(READING, 49)
    supplierIBAN = Trim(String99(READING, 293))
    supplierBIC = Trim(String99(READING, 294))
    If numberFalse > 0 Then
        MsgBox "Er zijn minstens " & numberFalse & " verbeteringen uit te voeren in Setup Bedrijfsinformatie", vbExclamation
    End If
    If Dir(LOCATION_COMPANYDATA + "vpeSjbs\pdfDDEF120.Txt") = "" Then
        Beep
        Msg = LOCATION_COMPANYDATA & "vpeSjbs\" & vbCrLf
        Msg = Msg & "ontbreekt nederlandstalige factuurdefinitie (pdfDDEF120.txt)." & vbCrLf & vbCrLf
        Msg = Msg & "Genereer een voorbeeld via de knop Beheer van VPE Lay-Out" & vbCrLf
        Msg = Msg & "Kies voor Submenu Bestand/Openen"
        MsgBox Msg, vbCritical
    Else
        If numberFalse = 0 Then checkSetUp = True
    End If

End Function


Private Sub InstalKlant()
Dim Klantje As String
Dim T       As Integer
Dim ktrlKlant As Boolean

Set rsKlant = New ADODB.Recordset
On Error Resume Next
Err = 0
rsKlant.CursorLocation = adUseClient
Msg = "SELECT TOP 1 * FROM Klanten WHERE A110 = '" + XLogKey + "'"
SnelHelpPrint Msg, BL_LOGGING
rsKlant.Open Msg, adntDB, adOpenDynamic, adLockOptimistic
If Err Or rsKlant.RecordCount <> 1 Then
    MsgBox "Onverwachte situatie"
    Exit Sub
End If
   
VerkoopDetail.Enabled = True
Klantje = vbCrLf & RV(rsKlant, "A100") & vbCrLf & _
        RV(rsKlant, "A125") & vbCrLf & _
        RV(rsKlant, "A104") & vbCrLf & _
        RV(rsKlant, "A109") & " " & _
        RV(rsKlant, "A107") & " " & RV(rsKlant, "A108")

customerVatNumber = Trim(RV(rsKlant, "A161"))
customerCompanyId = Trim(RV(rsKlant, "v404"))
customerA110 = Trim(RV(rsKlant, "A110"))

If (customerVatNumber + customerCompanyId) = "" Then
    Me.cbCheckTools.Enabled = False
    Me.cbUncl1001.Visible = False
ElseIf Len(customerCompanyId) > 0 And Len(customerVatNumber) = 0 Then
    Msg = "Een Belgisch ondernemingsnummer (" & customerCompanyId & ") en geen btwnummer" & vbCrLf & vbCrLf
    Msg = Msg & "Overtuig U ervan dat de klant geen btw plichtige is." & vbCrLf & vbCrLf
    Msg = Msg & "Hierna wordt verder gewerkt als gewone klant relatie zonder Peppol verplichting"
    MsgBox Msg, vbExclamation
    Me.cbCheckTools.Enabled = False
    Me.cbUncl1001.Visible = False
Else
    If VerkoopOptie(0).Value = True Or CreditNota.Value = True Then
        If customerCompanyId = "" Then
        Else
            ktrlKlant = Me.checkForB2BInvoice
            'If ktrlKlant = False Then
            '    Schoon
            '    Exit Sub
            'End If
        End If
    End If
End If

If customerCompanyId = "" Then
ElseIf customerCountryCode = "BE" Then
    Dim beCustomerPeppolCheck As Boolean
    beCustomerPeppolCheck = CheckCustomerDocuments(customerA110)
    If beCustomerPeppolCheck = False Then
        Msg = "Deze BE Klant is niet Peppol klaar" & vbCrLf & vbCrLf
        Msg = Msg & "In principe moet je deze klant als particulier bedienen met verkoopdocumenten. Controleer online ondernemingsnummer en btw nummer" & vbCrLf & vbCrLf
        Msg = Msg & "In principe mag je deze klant geen B2B verkoopdocumenten afleveren."
        MsgBox Msg, vbInformation
    End If
End If

Schoonvegen.Enabled = True
Sjabloon.Enabled = True

    DagKoers.Caption = "Dagkoers " & RV(rsKlant, "vs03")
    sMuntKlant = RV(rsKlant, "vs03")

    If Trim(RV(rsKlant, "v255")) = "" Then
        isDOM = False
        sepaDOM.Value = vbUnchecked
    Else
        isDOM = True
        sepaDOM.Value = vbChecked
        Overschrijvingsstrook.Value = vbChecked
    End If
    
    Set rsDetail = New ADODB.Recordset
    On Error Resume Next
    Err = 0
    rsDetail.CursorLocation = adUseClient
    Msg = "SELECT TOP 1 * FROM Allerlei WHERE v005 = '10" + sMuntKlant + "'"
    SnelHelpPrint Msg, BL_LOGGING
    rsDetail.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
    'cmdSwitch.Enabled = True
    If rsDetail.RecordCount <> 1 Then
        MsgBox "Muntkode eerst inbrengen via diverse gebruikersfiches a.u.b.  Nu wordt automatisch verdergewerkt met BEF", vbInformation, "Muntkode is " + sMunt
        If bhEuro Then
            DagKoers.Caption = "Munt in EUR"
            sMuntKlant = "EUR"
            dMuntK = 1
        Else
            DagKoers.Caption = "Munt in BEF"
            sMuntKlant = "BEF"
            dMuntK = 1
        End If
    Else
        TekstInfo5.Text = Dec$(Val(ADOBIB_TEXT(rsDetail("Memo"), "#v040 #")), "###.########")
        dMuntK = Val(ADOBIB_TEXT(rsDetail("Memo"), "#v040 #"))
    End If

    If bhEuro Then
        If sMuntKlant = "BEF" Then
            cmdSwitch.Caption = "Ingave in BEF"
            'cmdSwitch.Enabled = True
            dMuntK = 1
            'TekstInfo5.Text = Dec$(1 / EURO, "##0.########")
        ElseIf sMuntKlant = "EUR" Then
            cmdSwitch.Caption = "Ingave in EUR"
            'cmdSwitch.Enabled = True
            dMuntK = 1
        Else
            cmdSwitch.Caption = "Ingave in EUR"
            'cmdSwitch.Enabled = False
        End If
    Else
        If sMuntKlant = "EUR" Then
            cmdSwitch.Caption = "Ingave in EUR"
            'cmdSwitch.Enabled = True
            dMuntK = 1
        ElseIf sMuntKlant = "BEF" Then
            cmdSwitch.Caption = "Ingave in BEF"
            'cmdSwitch.Enabled = True
            dMuntK = 1
        Else
            cmdSwitch.Caption = "Ingave in BEF"
            'cmdSwitch.Enabled = False
        End If
    End If
    If sMuntKlant = "BEF" Or sMuntKlant = "EUR" Then
        TekstInfo5.Text = Dec$(dMuntK, "##0.########")
    Else
        dMuntK = Val(TekstInfo5.Text)
    End If
    DIRECTSELL_STRING = cmdSwitch.Caption
    rsDetail.Close
    Set rsDetail = Nothing
    
    If RV(rsKlant, "v149") = "" Then
        MsgBox "Landnummer is verplicht !"
        Exit Sub
    ElseIf RV(rsKlant, "v149") = "002" Then
        KlantInfo.Caption = vSet(RV(rsKlant, "A110"), 12) + "* Binnenland * " + Klantje
        VerkoopFLG = 0
        Medekontraktant.Enabled = True
        Dim btwBE As String
        btwBE = Trim(RV(rsKlant, "A161"))
        If Len(btwBE) = 10 Then
            If BtwKontrole(btwBE, False) = "" Then
                MsgBox "Btwnummer " & btwBE & " is ongeldig. Verbeter a.u.b. in klantfiche!", vbCritical
                Schoon
                Exit Sub
            End If
        End If
        'kontroleren LU speciaal
    ElseIf InStr(SISO, RV(rsKlant, "v149")) Then
        KlantInfo.Caption = vSet(RV(rsKlant, "A110"), 12) + "* E.U. mét Btw-nummer * " + Klantje
        VerkoopFLG = 1
        Medekontraktant.Enabled = False
        If vSet(RV(rsKlant, "A161"), 12) = Space$(12) Then
            KlantInfo.Caption = vSet(RV(rsKlant, "A110"), 12) + "* E.U. geen Btw-nummer * " + Klantje
            VerkoopFLG = 0
        End If
    Else
        KlantInfo.Caption = vSet(RV(rsKlant, "A110"), 12) + "* Uitvoer buiten E.U. *" + Klantje
        VerkoopFLG = 2
        Medekontraktant.Enabled = False
    End If

    TekstInfo0.Enabled = True
    TekstInfo1.Enabled = True
    TekstInfo2.Enabled = True

    If RV(rsKlant, "v151") = "1" Then
        Medekontraktant.Value = 1
    End If
    
    If Left(RV(rsKlant, "v161"), 3) = "400" Then
        TekstInfo2.Text = RV(rsKlant, "v161")
        KlantRekening = RV(rsKlant, "v161")
        If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", KlantRekening) Then
            Beep
            KlantRekening = String99(READING, 9)
            TekstInfo2.Text = KlantRekening
        End If
    Else
        KlantRekening = String99(READING, 9)
        TekstInfo2.Text = KlantRekening
    End If
    If Left(RV(rsKlant, "v225"), 2) = "70" Then
        DefaultVerkoop = RV(rsKlant, "v225")
        If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", DefaultVerkoop) Then
            Beep
            DefaultVerkoop = String99(READING, 25)
        End If
    Else
        DefaultVerkoop = String99(READING, 25)
    End If
    
    TekstInfo1.Text = VValdag((TekstInfo0.Text), RV(rsKlant, "vs04"))
    IsErKlassement
    If RV(rsKlant, "v253") <> "" Then
        cmdSQLInfo.Visible = True
    End If

    If Trim(RV(rsKlant, "g101")) = "0" Then
        customerPrefersEmail = False
    ElseIf Trim(RV(rsKlant, "v224")) <> "" Then
        customerPrefersEmail = True
    Else
        customerPrefersEmail = False
    End If
End Sub

Private Sub IsErKlassement()

On Error Resume Next
Set rsDetail = New ADODB.Recordset
Err = 0
rsDetail.CursorLocation = adUseClient
Msg = "SELECT TOP 1 * FROM Allerlei WHERE v004 = 'K" + RV(rsKlant, "A110") + "' AND v005 Like '" + dokumentType + "%'"
SnelHelpPrint Msg, BL_LOGGING
rsDetail.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Or rsDetail.RecordCount = 0 Then
    Klassement.Enabled = False
    Klassement.FontBold = False
Else
    Klassement.Enabled = True
    Klassement.FontBold = True
End If
rsDetail.Close
Set rsDetail = Nothing

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If Mim.Report.IsOpen = True Then
        Mim.Report.CloseDoc
   End If

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

frmDokHistoriek.lstDokHistoriek.Clear
frmDokHistoriek.Hide
If bhEuro Then
    DIRECTSELL_STRING = "Ingave in EUR"
Else
    DIRECTSELL_STRING = "Ingave in BEF"
End If
    
End Sub

Private Sub KlantAktiveren_Click()

Schoon

If VerkoopDetail.ListCount Then
    Msg = "Huidige inbreng en klant negeren." + vbCrLf + vbCrLf
    Msg = Msg + "Bent U zeker."
    KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, "Andere klant kiezen")
    If KtrlBox = vbYes Then
    Else
        Exit Sub
    End If
End If

aIndex = 1
SharedFl = TABLE_CUSTOMERS

GridText = ""
SqlSearch.Show 1
If Ktrl = 0 Then
    InstalKlant
Else
    Schoon
End If

End Sub

Private Sub Klassement_Click()
Dim aa As Variant
Dim T As Integer
Dim LaatsteWAS As String
Dim TotaalEX As Currency

If KlantInfo.Caption = "" Then Exit Sub
Unload Xlog
Xlog.X.Rows = 1
Xlog.X.Cols = 5

aa = ""
T = 0

Set rsDetail = New ADODB.Recordset
    
On Error Resume Next
Err = 0
rsDetail.CursorLocation = adUseClient
Msg = "SELECT * FROM Allerlei WHERE v004 = 'K" + RV(rsKlant, "A110") + "' AND v005 Like '" + dokumentType + "%' ORDER BY v004, v005 DESC"
SnelHelpPrint Msg, BL_LOGGING
Screen.MousePointer = vbHourglass
rsDetail.Open Msg, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Screen.MousePointer = vbNormal
    Exit Sub
ElseIf rsDetail.RecordCount = 0 Then
    MsgBox "Geen dokumenten (meer) te vinden.", vbInformation
    Screen.MousePointer = vbNormal
    Exit Sub
Else
    Do While Not rsDetail.EOF
        GoSub VolgendeLijn
        rsDetail.MoveNext
    Loop
End If
Screen.MousePointer = vbNormal

Load Xlog
Xlog.Caption = "dokument inladen voor : " + RV(rsKlant, "A100")
Xlog.X.Col = 0
Xlog.X.Row = 0
Xlog.X.Text = "dokument"
Xlog.X.Col = 1
Xlog.X.Text = "Datum"
Xlog.X.Col = 2
Xlog.X.Text = "Bedrag EUR"
Xlog.X.Col = 3
Xlog.X.Text = "Bedrag BEF"
Xlog.X.Col = 4
Xlog.X.Text = "Zie ook dokument"
Xlog.X.Row = 1
Xlog.X.Col = 0

Xlog.X.ColWidth(0) = 1170
Xlog.X.ColWidth(1) = 1020
Xlog.X.ColWidth(2) = 1020
Xlog.X.ColWidth(3) = 1020
Xlog.X.ColWidth(4) = 4560
Xlog.WijzigenLijn.Visible = False
XLogKey = ""
rsDetail.Close
Set rsDetail = Nothing
Xlog.SSTab1.TabVisible(1) = False
Xlog.Show 1
Unload Xlog
If XLogKey <> "" Then
    LaadHetdokument
End If
Exit Sub

VolgendeLijn:
If LaatsteWAS = ADOBIB_TEXT(rsDetail("Memo"), "#v033 #") Then
Else
    LaatsteWAS = ADOBIB_TEXT(rsDetail("Memo"), "#v033 #")
    aa = ADOBIB_TEXT(rsDetail("Memo"), "#v033 #") & vbTab
    aa = aa & DATE_TEXT(ADOBIB_TEXT(rsDetail("Memo"), "#v035 #")) & vbTab
    
    If ADOBIB_TEXT(rsDetail("Memo"), "#vEUR #") = "EUR" Then
        aa = aa & Format(CDbl(ADOBIB_TEXT(rsDetail("Memo"), "#v137 #")), "#,##0.00") & vbTab
        aa = aa & Format(CDbl(ADOBIB_TEXT(rsDetail("Memo"), "#v137 #")) * EURO, "#,##0.00") & vbTab
    Else
        aa = aa & Format(CDbl(ADOBIB_TEXT(rsDetail("Memo"), "#v137 #")) / EURO, "#,##0.00") & vbTab
        aa = aa & Format(CDbl(ADOBIB_TEXT(rsDetail("Memo"), "#v137 #")), "#,##0.00") & vbTab
    End If
    aa = aa & ADOBIB_TEXT(rsDetail("Memo"), "#v147 #") + " " + ADOBIB_TEXT(rsDetail("Memo"), "#v148 #")
    Xlog.X.AddItem aa
End If
Return

End Sub

Private Sub MaakTotaal()
Dim BTWIn As Double
Dim BTWEx As Double

Dim BtwEuroIn As Double
Dim BtwEuroEx As Double

Dim Teller As Integer

On Local Error Resume Next

For Teller = 1 To 3
    BTWBedrag(Teller) = Val(Dec((BTWBasis(Teller)) * Val(Mid(fmarBoxText("002", "2", Format(Teller)), 4, 4)) / 100, MASK_EURBH))
    BTWIn = BTWIn + BTWBasis(Teller) + BTWBedrag(Teller)
    BTWEx = BTWEx + BTWBasis(Teller)

    BTWEuroBedrag(Teller) = Val(Dec(BTWEuroBasis(Teller) * Val(Mid(fmarBoxText("002", "2", Format(Teller)), 4, 4)) / 100, MASK_EURBH))
    BtwEuroIn = BtwEuroIn + BTWEuroBasis(Teller) + BTWEuroBedrag(Teller)
    BtwEuroEx = BtwEuroEx + BTWEuroBasis(Teller)
Next
If cmdSwitch.Caption = "Ingave in BEF" Then
    lblBTWBedrag1 = Format(Round(BTWBasis(1)), "#,##0.00")
    lblBTWBedrag2 = Format(Round(BTWBasis(2)), "#,##0.00")
    lblBTWBedrag3 = Format(Round(BTWBasis(3)), "#,##0.00")
Else
    lblBTWBedrag1 = Format(BTWEuroBasis(1), "#,##0.00")
    lblBTWBedrag2 = Format(BTWEuroBasis(2), "#,##0.00")
    lblBTWBedrag3 = Format(BTWEuroBasis(2), "#,##0.00")
End If

BTWEx = BTWEx + BTWBasis(0)
BTWIn = BTWIn + BTWBasis(0)

BtwEuroEx = BtwEuroEx + BTWEuroBasis(0)
BtwEuroIn = BtwEuroIn + BTWEuroBasis(0)

LblExBtw = Format(Round(BTWEx, 2), "#,##0.00")   'in BEF
LblInBtw = Format(Round(BTWIn), "#,##0.00")    'in BEF
LblEx2Btw = Format(BtwEuroEx, "#,##0.00") 'masker voor EURO
LblIn2Btw = Format(BtwEuroIn, "#,##0.00")  'masker voor EURO

End Sub

Private Sub Medekontraktant_Click()

    If Medekontraktant.Value = vbChecked Then
        Msg = "Behoud alle BTW percentages a.u.b." & vbCr & vbCr
        Msg = Msg & "NOOIT via VAK 00 voor 0% foceren" & vbCr
        Msg = Msg & "Het programma doet zelf wel automatisch het nodige"
        MsgBox Msg, vbExclamation
        LblInBtw.Visible = False
        LblIn2Btw.Visible = False
    Else
        LblInBtw.Visible = True
        LblIn2Btw.Visible = True
    End If
    RefreshBTW

End Sub

Private Sub Optimaliseer_Click()
Dim HuidigProdukt As String * 13
Dim TempoS As String
Dim TT As Integer
Dim TTT As Integer
Dim TTTT As Integer
Dim TTTTT As Integer


If VerkoopDetail.ListCount = 0 Then Exit Sub

      Msg = "Meermaals voorkomende produktnummers samenvoegen tot één." + vbCrLf + vbCrLf
Msg = Msg + "Bent U zeker ?"
KtrlBox = MsgBox(Msg, 292)
If KtrlBox = 6 Then
    TT = 0
    Do While TT < VerkoopDetail.ListCount
        If Right(VerkoopDetail.List(TT), 1) = "0" Then
            HuidigProdukt = Mid(VerkoopDetail.List(TT), 98, 13)
            TTTTT = TT
            TTTT = TT + 1
            Do While TTTT < VerkoopDetail.ListCount
                If HuidigProdukt = Mid(VerkoopDetail.List(TTTT), 98, 13) Then
                    TempoS = VerkoopDetail.List(TTTTT)
                    Mid(TempoS, 54, 7) = Dec$(Val(Mid(VerkoopDetail.List(TTTT), 54, 7)) + Val(Mid(VerkoopDetail.List(TTTTT), 54, 7)), "####.00")
                    Mid(TempoS, 62, 12) = Dec$(Val(Mid(TempoS, 42, 11)) * Val(Mid(TempoS, 54, 7)) * Val(Mid(TempoS, 75, 6)), "#########.00")
                    VerkoopDetail.RemoveItem TTTT
                    VerkoopDetail.List(TTTTT) = TempoS
                    TTT = TTT + 1
                End If
                TTTT = TTTT + 1
            Loop
        End If
        TT = TT + 1
    Loop
    MsgBox "Einde optimalisatie.  " + Format(TTT) + " lijnen samengevoegd."
    RefreshBTW
End If

End Sub

Private Sub PrintUserDef(TypeEnTaal As String)
Dim T               As Integer
Dim Tabulatie       As Integer
Dim FlFree          As Integer
Dim Teller          As Integer

ReDim psTekst(50) As String
Dim Dummy As String

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
    If psY(PsLokatie) > 10500 Then
        Printer.CurrentY = psY(PsLokatie) - OVSStrooklijnen
    Else
        Printer.CurrentY = psY(PsLokatie)
    End If
    Printer.Print psTekst(PsLokatie);
    PsLokatie = PsLokatie + 1
Loop
For T = 0 To 10
    If BeginX(T) = 0 Then
    ElseIf Box(T) Then
        Printer.Line (BeginX(T), BeginY(T))-(TotX(T), TotY(T)), , B
    Else
        'Printer.Line (BeginX(T), BeginY(T))-(TotX(T), TotY(T))
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
            If FiguurY + Hoogte > 10500 Then
                Printer.PaintPicture Mim.imgFiguur.Picture, FiguurX, FiguurY - OVSStrooklijnen, Breedte, Hoogte
            Else
                Printer.PaintPicture Mim.imgFiguur.Picture, FiguurX, FiguurY, Breedte, Hoogte
            End If
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


Private Sub RasterSchoon()
Dim T As Integer

VerkoopDetail.Clear
GridText = ""
For T = 0 To 3
    BTWBasis(T) = 0
    BTWBedrag(T) = 0
    
    BTWEuroBasis(T) = 0
    BTWEuroBedrag(T) = 0
Next
TotaalBTW = 0
TotaalUitvoer = 0
MaakTotaal

End Sub

Private Sub RefreshBTW()
Dim aa As String
Dim T As Integer
Dim bVak As Integer
Dim maskerMULTI As String

maskerMULTI = MASK_EURBH
For T = 0 To 3
    BTWBasis(T) = 0
    BTWBedrag(T) = 0
    BTWEuroBasis(T) = 0
    BTWEuroBedrag(T) = 0
Next
TotaalBTW = 0
TotaalUitvoer = 0

For T = 0 To VerkoopDetail.ListCount - 1
    aa = VerkoopDetail.List(T)
    Select Case Right(aa, 1)
        Case "0", "1"
            bVak = Val(Mid(aa, 88, 1))
            If bVak = 6 Then
                If cmdSwitch.Caption = "Ingave in EUR" Then
                    BTWBasis(0) = BTWBasis(0) + Val(Dec(Val(Mid(aa, 62, 12)) * EURO, MASK_EURBH))
                    BTWEuroBasis(0) = BTWEuroBasis(0) + Val(Dec(Val(Mid(aa, 62, 12)), MASK_EURBH))
                Else
                    BTWBasis(0) = BTWBasis(0) + Val(Dec(Val(Mid(aa, 62, 12)), MASK_EURBH))
                    BTWEuroBasis(0) = BTWEuroBasis(0) + Val(Dec(Val(Mid(aa, 62, 12)) / EURO, MASK_EURBH))
                End If
            Else
                If cmdSwitch.Caption = "Ingave in EUR" Then
                    If BTWBasis(bVak) <> 0 Then
                        BTWBasis(bVak) = Round(BTWBasis(bVak))
                        maskerMULTI = "############"
                    End If
                    BTWBasis(bVak) = BTWBasis(bVak) + Val(Dec(Val(Mid(aa, 62, 12)) * EURO, maskerMULTI))
                    BTWEuroBasis(bVak) = BTWEuroBasis(bVak) + Val(Dec(Val(Mid(aa, 62, 12)), MASK_EURBH))
                Else
                    If BTWBasis(bVak) <> 0 Then
                        maskerMULTI = "############"
                        BTWBasis(bVak) = Val(Dec((BTWBasis(bVak)), maskerMULTI))
                    End If
                    BTWBasis(bVak) = BTWBasis(Val(Mid(aa, 88, 1))) + Val(Dec(Val(Mid(aa, 62, 12)), maskerMULTI))
                    BTWEuroBasis(bVak) = BTWEuroBasis(Val(Mid(aa, 88, 1))) + Val(Dec(Val(Mid(aa, 62, 12)) / EURO, MASK_EURBH))
                End If
            End If
    End Select
Next
MaakTotaal

End Sub

Private Function Schoon()

reprintOnly = False

Dim T As Integer

orderMarReferences = ""
Me.ButtonInfoSupported.Enabled = False
Me.ButtonInfoSupported.Visible = False
orderMarReferences = ""
Me.cbUncl1001.Visible = False
Me.cbUncl1001Toggle.Visible = False
Me.FramePeppol.Visible = False
Me.OptionPEPPOL_V3.Visible = False
Me.OptionUBL_BE_3_0.Visible = False

TLB_RECORD(TABLE_VARIOUS) = ""
dokumentHistoriek = ""
VerkoopDetail.Enabled = False
Afsluiten.Enabled = False
Me.CmbExtraAfdruk.Enabled = False
cmdSwitch.Enabled = False
Medekontraktant.Enabled = True
Sjabloon.Enabled = False
Klassement.Enabled = False
Klassement.FontBold = False
chkBTWBouw.Value = 0
VAT_BOBTHEBUILDERS = False
KlantInfo.Caption = ""
Annuleren.Enabled = True
Medekontraktant.Value = 0
If VerkoopOptie(0).Enabled = True Then CreditNota.Value = 0
customerVatNumber = ""

Err = 0
On Error Resume Next
TekstInfo0.Text = MIM_GLOBAL_DATE
If Err Then MsgBox "Landinstellingen voor België voorzien a.u.b.  Het programma wordt hierna beëindigd.", vbCritical: End
   
TekstInfo1.Text = MIM_GLOBAL_DATE
TekstInfo2.Text = KlantRekening
LblExBtw = ""
LblInBtw = ""
LblEx2Btw = ""
LblIn2Btw = ""
If Me.lstKopiePlak.ListCount Then
    Msg = "Gekopiëerde lijnen behouden ?"
    Ktrl = MsgBox(Msg, vbYesNo + vbDefaultButton1)
    Select Case Ktrl
        Case vbNo
            Me.lstKopiePlak.Clear
            frmDokHistoriek.lstDokHistoriek.Clear
    End Select
End If
frmDokHistoriek.Hide
           
RasterSchoon
'VerkoopOptie(0).Value = 1
Select Case Vr
    Case 11
        VerkoopOptie_Click 0
    Case 13
        CreditNota_Click
    Case 73
        VerkoopOptie_Click 1
    Case 59
        VerkoopOptie_Click 2
    Case Else
        MsgBox "Stop"
End Select
cmdSQLInfo.Visible = False
customerPrefersEmail = False
If String99(READING, 21) = "1" Then
'    MsgBox "Outlook Email wordt niet meer gebruikt in marIntegraal versie 10.4 of hoger.  Schakel deze instelling uit en gebruik VPE/PDF en/of de mario mailspooler eens alle toestellen beschikken over marIntegraal 10.4 of hoger", vbInformation
    useEmail = True
    On Local Error Resume Next
    Me.MPISessie.SignOff
    On Local Error Resume Next
    Me.MPISessie.SignOn
    If Err Then
        MsgBox Err.Description
    Else
        On Local Error GoTo 0
        Me.MPIBericht.SessionID = Me.MPISessie.SessionID
        SnelHelpPrint "E-mail sessie met succes opgestart. IDkode :" + Format(Me.MPISessie.SessionID), BL_LOGGING
    End If
End If

On Error Resume Next
DirekteVerkoop.SetFocus

End Function

Private Sub SchoonVegen_Click()

Schoon

End Sub

Private Sub Sjabloon_Click()
Dim BsTekst As String * 40
Dim T As Integer
Dim LeftPos As Integer

GridText = ""
For T = 1 To VerkoopDetail.ListCount
    VerkoopDetail.ListIndex = T - 1
    GridText = GridText + VerkoopDetail.Text + Chr$(13)
Next
SjabloonIO.Show 1
If GridText <> "" Then
    Do While VerkoopDetail.ListCount
        VerkoopDetail.RemoveItem 0
    Loop
    GridText = GridText + Chr$(13)
    Do While GridText <> ""
        LeftPos = InStr(GridText, Chr$(13))
        If LeftPos <= 1 Then
            GridText = ""
            Exit Do
        End If
        VerkoopDetail.AddItem Left(GridText, InStr(GridText, Chr$(13)) - 1)
        GridText = Right(GridText, Len(GridText) - LeftPos)
    Loop
    RefreshBTW
    Afsluiten.Enabled = True
    Me.CmbExtraAfdruk.Enabled = True
End If

End Sub


Private Sub TekstInfo0_GotFocus()

    TekstInfo0.SelStart = 0
    TekstInfo0.SelLength = Len(TekstInfo0.Text)

End Sub

Private Sub TekstInfo1_GotFocus()

    TekstInfo1.SelStart = 0
    TekstInfo1.SelLength = Len(TekstInfo1.Text)

End Sub


Private Sub TekstInfo2_GotFocus()

    TekstInfo2.SelStart = 0
    TekstInfo2.SelLength = Len(TekstInfo2.Text)

End Sub


Private Sub TekstInfo3_GotFocus()

    TekstInfo3.SelStart = 0
    TekstInfo3.SelLength = Len(TekstInfo3.Text)

End Sub


Private Sub TekstInfo5_GotFocus()

    TekstInfo5.SelStart = 0
    TekstInfo5.SelLength = Len(TekstInfo5.Text)

End Sub

Private Sub TekstInfo0_LostFocus()

    On Local Error Resume Next

    If DATE_INVALID((TekstInfo0.Text)) Then
        TekstInfo0.Text = MIM_GLOBAL_DATE
        TekstInfo0.SetFocus
    ElseIf Not DATE_CHECK((TekstInfo0.Text), PERIODAS_TEXT) Then
        If dokumentType = "15" Then
            BJPERDAT.WindowState = 0
            BJPERDAT.PeriodeBoekjaar.SetFocus
        End If
    Else
        TekstInfo1.Text = VValdag((TekstInfo0.Text), RV(rsKlant, "vs04"))
    End If

End Sub

Private Sub TekstInfo1_LostFocus()

    On Local Error Resume Next

    If DATE_INVALID((TekstInfo1.Text)) Then
        TekstInfo1.Text = MIM_GLOBAL_DATE
        TekstInfo1.SetFocus
    End If

End Sub

Private Sub TekstInfo2_LostFocus()

    On Local Error Resume Next

    If Left(TekstInfo2.Text, 2) <> "40" Then
        TekstInfo2.Text = KlantRekening
        TekstInfo2.SetFocus
    End If
        
    If Not ADO_GET(TABLE_LEDGERACCOUNTS, 0, "=", TekstInfo2) Then
        TekstInfo2.Text = KlantRekening
        TekstInfo2.SetFocus
    End If
    
End Sub


Private Sub TekstInfo3_LostFocus()

    On Local Error Resume Next

    dokumentSleutel = TekstInfo3.Text
    DirekteVerkoop.Caption = vSet("Verkoopverrichting", 28) + "(" + dokumentSleutel + ")"

End Sub


Private Sub TekstInfo5_LostFocus()

    On Local Error Resume Next

    dMuntK = Val(TekstInfo5.Text)
    TekstInfo5.Text = Dec$(Val(TekstInfo5.Text), "##0.000000")
    
End Sub

Private Sub Verkoopdetail_DblClick()

If VerkoopDetail.ListIndex = -1 Then
    VerkoopDetail_KeyPress 43
Else
    VerkoopDetail_KeyPress 13
End If

End Sub

Private Sub VerkoopDetail_GotFocus()

If CmdStock.Default Then CmdStock.Default = False
If CmdOmschrijving.Default Then CmdOmschrijving.Default = False
If CmdTekst.Default Then CmdTekst.Default = False
KlantAktiveren.Default = False
SnelHelpPrint "[S]tock, [O]mschrijving, Vrije [T]ekst = lijn bijvoegen, [Insert] = plakken speciaal, [Delete] = knippen speciaal of verwijderen, [Enter] om te wijzigen", BL_LOGGING

End Sub

Private Sub VerkoopDetail_KeyDown(KeyCode As Integer, Shift As Integer)
Dim Positie As Integer
Dim TempoBedrag As Double
Dim TempoBTW As Double
Dim COUNT_TOHier As Integer
Dim StringHier As String

Select Case KeyCode
    Case 45     'Insert
        If Me.lstKopiePlak.ListCount = 0 Then
            MsgBox "Er zijn geen lijnen (meer) te plakken!"
        Else
            For COUNT_TOHier = 0 To Me.lstKopiePlak.ListCount - 1
                StringHier = Me.lstKopiePlak.List(COUNT_TOHier)
                VerkoopDetail.AddItem StringHier
                If Right(StringHier, 1) = "2" Then
                Else
                    RefreshBTW
                End If
            Next
            Me.lstKopiePlak.Clear
        End If

    Case 46     'Delete
        Positie = VerkoopDetail.ListIndex
        If Positie < 0 Then
            MsgBox "Eerst een lijn selekteren !", 0, "Lijn wijzigen"
            MaakTotaal
            Exit Sub
        End If
        If dokumentType <> "15" And Annuleren.Enabled = False Then
            Annuleren.Enabled = True
        End If
        GridText = VerkoopDetail.Text
        If Right(GridText, 1) = "2" Then
        Else
            RefreshBTW
        End If
        Msg = "Lijn bewaren om te plakken ?"
        Ktrl = MsgBox(Msg$, 36, "Lijn verwijderen/knippen ?")
        Select Case Ktrl
            Case 7
        Case Else
            Me.lstKopiePlak.AddItem VerkoopDetail.Text
        End Select
        VerkoopDetail.RemoveItem Positie
        RefreshBTW
        If VerkoopDetail.ListCount <> 0 Then VerkoopDetail.ListIndex = Positie - 1
        DirekteVerkoop.SetFocus
End Select

End Sub

Private Sub VerkoopDetail_KeyPress(KeyAscii As Integer)
Dim Positie As Integer
Dim TempoBedrag As Double
Dim TempoBTW As Integer

If KeyAscii = 43 Then Exit Sub    '+

Jumper:
XLogKey = DefaultVerkoop
Select Case UCase(Chr(KeyAscii))
    Case "S"
        aIndex = 0
    Case "O"
        aIndex = 1
    Case "T"
        aIndex = 2
End Select
Positie = VerkoopDetail.ListIndex
Select Case KeyAscii
    Case 13         'Enter
        If Positie < 0 Then
            MsgBox "Eerst een lijn selecteren !", 0, "Lijn wijzigen"
            Exit Sub
        End If
        If dokumentType <> "15" And Annuleren.Enabled = False Then
            Annuleren.Enabled = True
        End If
        GridText = VerkoopDetail.Text
        If Right(GridText, 1) = "2" Then
        Else
            RefreshBTW
        End If
        XLogKassa = ""
        WijzigenVerkoop.Show 1
        Unload WijzigenVerkoop
        DirekteVerkoop.SetFocus
        If GridText = "" Then
            RefreshBTW
        Else
            If Right(GridText, 1) = "2" Then
            Else
                TempoBedrag = Val(Mid(GridText, 62, 12))
                TempoBTW = Val(Mid(GridText, 88, 1))
                If TempoBTW = 6 Then
                    If bhEuro Then
                        If cmdSwitch.Caption = "Ingave in EUR" Then
                            BTWBasis(0) = BTWBasis(0) + TempoBedrag
                        Else
                            BTWBasis(0) = BTWBasis(0) + Round(TempoBedrag / EURO, 2)
                        End If
                    Else
                        If cmdSwitch.Caption = "Ingave in BEF" Then
                            BTWBasis(0) = BTWBasis(0) + TempoBedrag
                        Else
                            BTWBasis(0) = BTWBasis(0) + Round(TempoBedrag * EURO, 2)
                        End If
                    End If
                Else
                    If bhEuro Then
                        If cmdSwitch.Caption = "Ingave in EUR" Then
                            BTWBasis(TempoBTW) = BTWBasis(TempoBTW) + TempoBedrag
                        Else
                            BTWBasis(TempoBTW) = BTWBasis(TempoBTW) + Round(TempoBedrag / EURO, 2)
                        End If
                    Else
                        If cmdSwitch.Caption = "Ingave in BEF" Then
                            BTWBasis(TempoBTW) = BTWBasis(TempoBTW) + TempoBedrag
                        Else
                            BTWBasis(TempoBTW) = BTWBasis(TempoBTW) + Round(TempoBedrag * EURO, 2)
                        End If
                    End If
                End If
            End If
            VerkoopDetail.RemoveItem Positie
            VerkoopDetail.AddItem GridText, Positie
        End If
        VerkoopDetail.ListIndex = Positie
        MaakTotaal

    Case 43
        GridText = ""
        XLogKassa = ""
        WijzigenVerkoop.Show 1
        Unload WijzigenVerkoop
        DirekteVerkoop.SetFocus
        If GridText = "" Then
        Else
            If Right(GridText, 1) = "2" Then
            Else
                TempoBedrag = Val(Mid(GridText, 62, 12))
                TempoBTW = Val(Mid(GridText, 88, 1))
                If TempoBTW = 6 Then
                    If bhEuro Then
                        If cmdSwitch.Caption = "Ingave in EUR" Then
                            BTWBasis(0) = BTWBasis(0) + TempoBedrag
                        Else
                            BTWBasis(0) = BTWBasis(0) + Round(TempoBedrag / EURO, 2)
                        End If
                    Else
                        If cmdSwitch.Caption = "Ingave in BEF" Then
                            BTWBasis(0) = BTWBasis(0) + TempoBedrag
                        Else
                            BTWBasis(0) = BTWBasis(0) + Round(TempoBedrag * EURO, 2)
                        End If
                    End If
                Else
                    If bhEuro Then
                        If cmdSwitch.Caption = "Ingave in EUR" Then
                            BTWBasis(TempoBTW) = BTWBasis(TempoBTW) + TempoBedrag
                        Else
                            BTWBasis(TempoBTW) = BTWBasis(TempoBTW) + Round(TempoBedrag / EURO, 2)
                        End If
                    Else
                        If cmdSwitch.Caption = "Ingave in BEF" Then
                            BTWBasis(TempoBTW) = BTWBasis(TempoBTW) + TempoBedrag
                        Else
                            BTWBasis(TempoBTW) = BTWBasis(TempoBTW) + Round(TempoBedrag * EURO, 2)
                        End If
                    End If
                End If
                MaakTotaal
            End If
            
            Dim splitmilieu() As String
            Dim telmilieu As Integer
            
            If Positie < 0 Then
                VerkoopDetail.AddItem GridText, VerkoopDetail.ListCount
                If blMilieu = True Then
                    splitmilieu = Split(MilieuGridText, vbCrLf)
                    For telmilieu = 0 To UBound(splitmilieu) - 1
                        VerkoopDetail.AddItem splitmilieu(telmilieu), VerkoopDetail.ListCount
                    Next
                    blMilieu = False
                End If
            Else
                VerkoopDetail.AddItem GridText, Positie
                If blMilieu = True Then
                    splitmilieu = Split(MilieuGridText, vbCrLf)
                    MsgBox "stop, waarom?"
                    For telmilieu = 0 To UBound(splitmilieu) - 1
                        VerkoopDetail.AddItem splitmilieu(telmilieu), VerkoopDetail.ListCount
                    Next
                    blMilieu = False
                End If
            End If
        End If
        If Mid(XLogKey, 39, 2) = vbCrLf Then
            Do While Len(XLogKey) <> 0
                Msg = Mid(XLogKey, 1, 38) + String(37, " ") + "|2"
                VerkoopDetail.AddItem Msg
                XLogKey = Mid(XLogKey, 41)
            Loop
        End If
        
    Case 79, 83, 84, 111, 115, 116
        If dokumentType <> "15" And Annuleren.Enabled = False Then
            Annuleren.Enabled = True
        End If
        'LijnType(InStr$("SOT", UCase$(Chr$(KeyAscii))) - 1).Value = 1
        KeyAscii = 43
        GoTo Jumper
    Case Else
'        VensterKeyPress KeyAscii
End Select
If VerkoopDetail.ListCount Then
    Afsluiten.Enabled = True
    Me.CmbExtraAfdruk.Enabled = True
Else
    Afsluiten.Enabled = False
    Me.CmbExtraAfdruk = False
End If
RefreshBTW
If CmdStock.Default Then CmdStock.SetFocus
If CmdOmschrijving.Default Then CmdOmschrijving.SetFocus
If CmdTekst.Default Then CmdTekst.SetFocus

End Sub

Private Sub VerkoopDetail_LostFocus()

VerkoopDetail.ListIndex = -1

End Sub

Private Sub VerkoopOptie_Click(Index As Integer)

Dim ktrlKlant As Boolean

If KlantInfo.Caption <> "" Then
    Select Case Index
        Case 0
            If customerCompanyId = "" Then
            Else
                ktrlKlant = Me.checkForB2BInvoice
            End If
            
        Case 1, 2

            If customerCompanyId = "" Then
            Else
                ktrlKlant = Me.checkForB2BInvoice
            End If
        
        Case Else
            
    
    End Select
End If

If dokumentHistoriek <> Space$(11) Then
    MsgBox "Opgelet, data van o.a. " + dokumentHistoriek + " is nog actief.  Vermijd dubbele bewerkingen.", 64
End If

If VerkoopDetail.ListCount Then
    If Annuleren.Enabled = False Then
        MsgBox "U gaat naar een hogere modus met ingeladen dokument(en)." + vbCrLf + vbCrLf + "Indien dit niet de bedoeling was, onmiddellijk verkoopvenster sluiten en herbeginnen a.u.b.", vbOKOnly + vbExclamation
        dokumentHistoriek = dokumentSleutel
        TLB_RECORD(TABLE_VARIOUS) = ""
        Annuleren.Enabled = True
    End If
End If

If Index > 0 Then
    CreditNota.Enabled = False
Else
    CreditNota.Enabled = True
End If

Select Case Index
    Case 0
        dokumentType = "15" 'Faktuur of creditnota
        CreditNota.Value = 0
        Vr = 11
        dokumentSleutel = SleutelDok(11)
        TekstInfo0.Text = MIM_GLOBAL_DATE
        TekstInfo1.Text = MIM_GLOBAL_DATE
        DirekteVerkoop.Caption = vSet("Verkoopverrichting", 28) + "(" + dokumentSleutel + ")"
        chkZonderRelatie.Visible = False
        
    Case 1
        dokumentType = "14" 'Bestelbon, Leveringsbon
        dokumentSleutel = SleutelDok(73)
        Vr = 73
        DirekteVerkoop.Caption = vSet("Bestelling/levering", 28) + "(" + dokumentSleutel + ")"
        chkZonderRelatie.Visible = True
        
    Case 2
        dokumentType = "13" 'Offerte
        dokumentSleutel = SleutelDok(59)
        Vr = 59
        DirekteVerkoop.Caption = vSet("Offerte", 28) + "(" + dokumentSleutel + ")"
        chkZonderRelatie.Visible = True
        
End Select
TekstInfo3 = dokumentSleutel
If Ktrl = 99 Then
          Msg = VerkoopOptie(Index).Caption + " actief bij andere gebruiker." + vbCrLf + vbCrLf
    Msg = Msg + "Verkoopverrichting afsluiten of andere optie selecteren a.u.b. !"
    Ktrl = MsgBox(Msg, 16)
    If Index = 0 Then VerkoopOptie(1).Value = 1
    If Index = 1 Then VerkoopOptie(0).Value = 1
    If Index = 2 Then VerkoopOptie(1).Value = 1
    VerkoopOptie(Index).Enabled = False
    Exit Sub
End If
IsErKlassement

End Sub

Private Function WegBoekFout%()
Dim T               As Integer
Dim AantalStuks     As Single

'Gewijzigd voor scanning
WegBoekFout = False
dMuntK = Val(TekstInfo5.Text)
DKTRL_CUMUL = 0: DKTRL_BEF = 0: DKTRL_EUR = 0
Unload frmBoeking
frmBoeking.Hide
TLB_RECORD(TABLE_INVOICES) = ""
'bijgevoegd voor scanning..
If rsMAR(TABLE_INVOICES).State = adStateClosed Then
    Ktrl = bOpen(TABLE_INVOICES)
End If
rsMAR(TABLE_INVOICES).AddNew

vBib TABLE_INVOICES, dokumentSleutel, "v033"
vBib TABLE_INVOICES, "K" + RV(rsKlant, "A110"), "v034"
vBib TABLE_INVOICES, DATE_KEY(TekstInfo0), "v035"
vBib TABLE_INVOICES, DATE_KEY((TekstInfo1.Text)), "v036"
vBib TABLE_INVOICES, Dec$(dMuntK, "###.0000000000"), "v040"
vBib TABLE_INVOICES, sMuntKlant, "vs03"

If CreditNota.Value = 0 Then
    If BTWBasis(0) <> 0 Then
        If bhEuro Then
            vBib TABLE_INVOICES, Str$(BTWEuroBasis(0)), "v055"
        Else
            vBib TABLE_INVOICES, Dec((BTWBasis(0)), MASK_SY(0)), "v055"
        End If
    End If
End If
For T = 1 To 3
    If bhEuro Then
        If BTWEuroBedrag(T) <> 0 Then
            If CreditNota.Value = 0 Then
                vBib TABLE_INVOICES, Str$(BTWEuroBasis(T)), "v" + Dec(55 + T, "000")
            End If
        End If
    Else
        If BTWBedrag(T) <> 0 Then
            If CreditNota.Value = 0 Then
                vBib TABLE_INVOICES, Dec((BTWBasis(T)), MASK_SY(0)), "v" + Dec(55 + T, "000")
            End If
        End If
    End If
Next

If bhEuro Then
    vBib TABLE_INVOICES, Str$(TotaalBTW), "v064"
Else
    vBib TABLE_INVOICES, Dec((TotaalBTW), MASK_SY(0)), "v064"
End If
If Vr = 11 Then
    If Medekontraktant.Value Then
        If bhEuro Then
            vBib TABLE_INVOICES, Str$(CDbl(LblEx2Btw)), "v059"
        Else
            vBib TABLE_INVOICES, Dec(CDbl(LblExBtw), MASK_SY(0)), "v059"
        End If
    End If
End If

If Vr = 11 And VerkoopFLG = 1 Then
    If bhEuro Then
        vBib TABLE_INVOICES, Str$(CDbl(LblEx2Btw)), "v060"
    Else
        vBib TABLE_INVOICES, Dec(CDbl(LblExBtw), MASK_SY(0)), "v060"
    End If
End If

If Vr = 11 And VerkoopFLG = 2 Then
    If bhEuro Then
        vBib TABLE_INVOICES, Str$(CDbl(LblEx2Btw)), "v061"
    Else
        vBib TABLE_INVOICES, Dec(CDbl(LblExBtw), MASK_SY(0)), "v061"
    End If
End If

If Vr = 13 And VerkoopFLG = 1 Then
    If bhEuro Then
        vBib TABLE_INVOICES, Str$(CDbl(LblEx2Btw)), "v062"
    Else
        vBib TABLE_INVOICES, Dec(CDbl(LblExBtw), MASK_SY(0)), "v062"
    End If
ElseIf Vr = 13 Then
    If bhEuro Then
        vBib TABLE_INVOICES, Str$(CDbl(LblEx2Btw)), "v063"
    Else
        vBib TABLE_INVOICES, Dec(CDbl(LblExBtw), MASK_SY(0)), "v063"
    End If
End If

If ForFait Then
    GoTo JumpTodokument
End If

For T = 0 To VerkoopDetail.ListCount - 1
    VerkoopDetail.ListIndex = T
    Select Case Right(VerkoopDetail.Text, 1)
        Case "0", "1"
            GoSub JournaalNew
            '1, 40      Omschrijving
            '42, 11     BedragGoed
            '54, 7      Aantal
            '62, 12     Totaal
            '75, 6      Verpakking
            '90, 7      VerkoopRekening
            '98, 13     ProduktNummer
            
            rsJournaal("v067") = Trim(Mid(VerkoopDetail.Text, 1, 30))
            If Trim(Mid(VerkoopDetail.Text, 98, 13)) = "OMSCHRIJVING" Then
                'niks
            Else
                rsJournaal("v102") = Trim(Mid(VerkoopDetail.Text, 98, 13))
            End If
            If Vr = 11 Then
                rsJournaal("v068") = Dec(-Val(Mid(VerkoopDetail.Text, 62, 12)), "#.00")
                rsJournaal("dece068") = Val(Dec(-Val(Mid(VerkoopDetail.Text, 62, 12)), "#.00"))
            Else
                rsJournaal("v068") = Dec(Val(Mid(VerkoopDetail.Text, 62, 12)), "#.00")
                rsJournaal("dece068") = Val(Dec(Val(Mid(VerkoopDetail.Text, 62, 12)), "#.00"))
            End If
            rsJournaal("v019") = Mid(VerkoopDetail.Text, 90, 7)
            If Not ADOJOURNAL_OK Then
                MsgBox Error
                Exit Function
            End If
            
            If Right(VerkoopDetail.Text, 1) = "0" And BJPERDAT.Boekjaar.ListIndex = 0 Then
                If Not ADO_GET(TABLE_PRODUCTS, 0, "=", Mid(VerkoopDetail.Text, 98, 13)) Then
                    Beep
                    MsgBox "stop"
                Else
                    AantalStuks = Val(Mid(VerkoopDetail.Text, 54, 7)) * Val(Mid(VerkoopDetail.Text, 75, 6))
                    If Vr = 11 Then
                        rsMAR(TABLE_PRODUCTS)("v120") = Dec(AantalStuks + Val(RV(rsMAR(TABLE_PRODUCTS), "v120")), MASK_SY(2))
                        rsMAR(TABLE_PRODUCTS)("e122") = Dec((AantalStuks * Val(RV(rsMAR(TABLE_PRODUCTS), "e113"))) + Val(RV(rsMAR(TABLE_PRODUCTS), "e122")), "##" & MASK_EURX)
                    Else
                        rsMAR(TABLE_PRODUCTS)("v120") = Dec(-AantalStuks + Val(RV(rsMAR(TABLE_PRODUCTS), "v120")), MASK_SY(2))
                        rsMAR(TABLE_PRODUCTS)("e122") = Dec((-AantalStuks * Val(RV(rsMAR(TABLE_PRODUCTS), "e113"))) + Val(RV(rsMAR(TABLE_PRODUCTS), "e122")), "##" & MASK_EURX)
                    End If
                    rsMAR(TABLE_PRODUCTS).Update
                End If
            End If
        Case Else
    End Select
Next

If TotaalBTW Then
    GoSub JournaalNew
    rsJournaal("v067") = RV(rsKlant, "A100")
    
    If Vr = 11 Then
        rsJournaal("v068") = Str$(-TotaalBTW)
        rsJournaal("dece068") = -TotaalBTW
        rsJournaal("v019") = rbtwVAK(0)
    Else
        rsJournaal("v068") = Str$(TotaalBTW)
        rsJournaal("dece068") = TotaalBTW
        rsJournaal("v019") = rbtwVAK(6)
    End If
    If Not ADOJOURNAL_OK Then
        MsgBox Error
        Exit Function
    End If
End If

GoSub JournaalNew
rsJournaal("v067") = RV(rsKlant, "A100")
rsJournaal("v019") = TekstInfo2.Text
If Vr <> 11 Then
    rsJournaal("v068") = Str$(-CDbl(LblEx2Btw) - TotaalBTW)
    rsJournaal("dece068") = -CDbl(LblEx2Btw) - TotaalBTW
Else
    rsJournaal("v068") = Str$(CDbl(LblEx2Btw) + TotaalBTW)
    rsJournaal("dece068") = CDbl(LblEx2Btw) + TotaalBTW
End If
If Not ADOJOURNAL_OK Then
    MsgBox Error
    Exit Function
End If

JumpTodokument:
If bhEuro Then
    vBib TABLE_INVOICES, Str$(CDbl(LblEx2Btw) + TotaalBTW), "v249"
Else
    If XisEuroWisBEF = False Then
        vBib TABLE_INVOICES, Dec(CDbl(LblExBtw) + TotaalBTW, MASK_SY(0)), "v249"
    Else
        vBib TABLE_INVOICES, Str$(Round((CDbl(LblExBtw) + TotaalBTW) / EURO, 2)), "v249"
    End If
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
    Select Case Left(Mim.cmdWegBoekModus.Text, 1)
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
Exit Function

FoutMelding:
MsgBox "Fout tijdens het wegboeken." + vbCrLf + "Hierna worden alle bewerkingen genegeerd." + vbCrLf + vbCrLf + "Begin opnieuw."
WegBoekFout = True
Exit Function

JournaalNew:
rsJournaal.AddNew
rsJournaal("v034") = "K" + RV(rsKlant, "A110")
rsJournaal("v066") = DATE_KEY(TekstInfo0)
rsJournaal("v035") = DATE_KEY(TekstInfo0)
rsJournaal("v033") = dokumentSleutel
rsJournaal("v069") = TekstInfo2.Text
Return

End Function

Function XmlInvoiceGenerateUBLBE30(thisPdf As String) As Boolean

'UBL
'base64Str = ""
documentTemplate = ""
invoiceLineTemplate = ""
invoiceTaxLineTemplate = ""
taxSubtotalLine = ""
listTaxSubtotalLines = ""
thisInvoiceLine = ""
listInvoiceLines = ""
tmpString = ""
'SUPPLIER:
'invoiceNumber = "" '240100225
'invoiceCustomerNumber = "" ' 844
invoiceCurrency = "" 'EUR
'invoiceDate = "" '2024-01-31
'invoiceDueDate = "" '2024-03-01
'invoiceTypeCode = "" '380 (see https://docs.peppol.eu/poac/pint/pint/trn-invoice/codelist/UNCL1001-inv/)
'orderReference = "" 'ex. 024010022505 as of +++024/0100/22505+++
'supplierTaxScheme = "" 'ex. VAT
'supplierRegistrationId = "" 'ex. 0423100736
'supplierVatNumber = "" 'ex. 0423100736 (Titeca)
'supplierRegistrationName = "" 'Titeca Accountancy Merelbeke NV
'supplierStreetName = "" 'Fraterstraat 132
'supplierCityName = "" 'MERELBEKE
'supplierPostalZone = "" '9820
'supplierCountryCode = "" 'BE
'supplierCompanyId = "" 'BE0423100736
'supplierTelephone = "" '09 232 28 00
'CUSTOMER:
'customerVatNumber = "" '0440058217
'customerTaxScheme = "" 'ex. VAT
'if customer is VAT then else ?
'customerName = "" 'ex. Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv
'customerStreetName = "" 'Grote Baan 141
'customerCityName = "" 'HERDERSEM
'customerPostalZone = "" '9310
'customerCountryCode = "" 'BE
'customerCompanyId = "" 'BE04400580217
'customerRegistrationName = "" 'Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv
'customerElectronicMail = "" 'info@rv.be
'PAYMENTMEANS/TERMS
paymentMeansCode = "1" '1 see: https://docs.peppol.eu/poac/pint/pint/trn-invoice/codelist/UNCL4461/
'paymentID = "" '024010022505
'supplierIBAN = "" 'BE50001683378618
'supplierBIC = "" 'GEBABEBB
'paymentTerms = "" '024010022505
'TAXTOTAL
subTaxTotalAmount = "" '435.86
subTaxableAmount = "" '2075.50
subTaxAmount = "" '435.86
subTaxCategory = "" '03
subTaxPercent = "" '21.00
subTaxExclusiveAmount = ""
subTaxInclusiveAmount = ""
subPayableAmount = ""

'taxGlobalTotalAmount = ""
'taxableGlobalTotalAmount = ""
globalTotalAmount = ""

'LEGALMTOTAL
taxExclusiveAmount = "" '2075.50
taxInclusiveAmount = "" '2511.36
payableAmount = "" '2511.36

'LINES
lineCounter = ""
lineQuantity = ""
lineExtentionAmount = ""
lineAccountingcost = ""
lineTaxAmountAccurate = ""
lineDescription = ""
lineSellerItemIdentification = ""
lineTaxCategory = ""
lineTaxPercentage = ""
linePriceAmount = ""
vatGroupsCount = 0
ublTT = 0

XmlInvoiceGenerateUBLBE30 = False
If Me.OptionUBL_BE_3_0.Enabled = False Then
    Exit Function
End If

If Left(invoiceNumber, 2) = "V0" Then
    Ktrl = ScrLeesBestandAlleTekst(documentTemplate, PROGRAM_LOCATION + "xml-templates\ubl_be_3_0-invoice.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
Else
    '                                                                                  ubl_be_3_0-creditnote
    Ktrl = ScrLeesBestandAlleTekst(documentTemplate, PROGRAM_LOCATION + "xml-templates\ubl_be_3_0-creditnote.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
End If

documentTemplate = Replace(documentTemplate, "{marVersion}", Trim(App.Major & "." & App.Minor & "." & App.Revision))

If Left(invoiceNumber, 2) = "V0" Then
    documentTemplate = Replace(documentTemplate, "{invoiceOrCreditnote}", "CommercialInvoice")
Else
    documentTemplate = Replace(documentTemplate, "{invoiceOrCreditnote}", "CreditNote")
End If

'invoiceNumber = "240100225"
documentTemplate = Replace(documentTemplate, "{documentNumber}", invoiceNumber)
'invoiceCustomerNumber = "844"
documentTemplate = Replace(documentTemplate, "{customerNumber}", invoiceCustomerNumber)
invoiceCurrency = "EUR"
documentTemplate = Replace(documentTemplate, "{currency}", invoiceCurrency)

'invoiceDate = "2024-01-31"
documentTemplate = Replace(documentTemplate, "{documentDate}", invoiceDate)
'invoiceDueDate = "2024-03-01"
If Left(invoiceNumber, 2) = "V0" Then
    documentTemplate = Replace(documentTemplate, "{documentDueDate}", invoiceDueDate)
End If
documentTemplate = Replace(documentTemplate, "{documentTypeCode}", invoiceTypeCode)
'orderReference = "024010022505" 'as of +++024/0100/22505+++

If orderMarReferences = "" Then
    documentTemplate = Replace(documentTemplate, "{orderReference}", invoiceNumber)
Else
    documentTemplate = Replace(documentTemplate, "{orderReference}", orderMarReferences)
End If

base64Str = GetBase64Str(thisPdf)
documentTemplate = Replace(documentTemplate, "{pdfFileName}", invoiceNumber + ".pdf")
documentTemplate = Replace(documentTemplate, "{pdfBase64}", base64Str)

documentTemplate = Replace(documentTemplate, "{supplierTaxScheme}", supplierTaxScheme)
'supplierVatNumber = "0423100736"
documentTemplate = Replace(documentTemplate, "{supplierVatNumber}", supplierVatNumber)
'tmpString = String99(READING, 46)
'If InStr(tmpString, "&") Then 'verbeteren voor XML bestand!!!
'    tmpString = Replace(tmpString, "&", "&amp;")
'End If
'supplierRegistrationId = "0423100736"
documentTemplate = Replace(documentTemplate, "{supplierRegistrationId}", supplierRegistrationId)
'supplierRegistrationName = "Titeca Accountancy Merelbeke NV"
documentTemplate = Replace(documentTemplate, "{supplierRegistrationName}", supplierRegistrationName)
'supplierStreetName = "Fraterstraat 132"
documentTemplate = Replace(documentTemplate, "{supplierStreetName}", supplierStreetName)
'supplierCityName = "MERELBEKE"
documentTemplate = Replace(documentTemplate, "{supplierCityName}", supplierCityName)
'supplierPostalZone = "9820"
documentTemplate = Replace(documentTemplate, "{supplierPostalZone}", supplierPostalZone)
'supplierCountryCode = "BE"
documentTemplate = Replace(documentTemplate, "{supplierCountryCode}", supplierCountryCode)
'supplierCompanyId = "0423100736"
documentTemplate = Replace(documentTemplate, "{supplierCompanyId}", supplierCompanyId)
'supplierCompanyIdExtended = "BE0423100736"
documentTemplate = Replace(documentTemplate, "{supplierCompanyIdExtended}", supplierCompanyIdExtended)

'supplierTelephone = "09 232 28 00"
documentTemplate = Replace(documentTemplate, "{supplierTelephone}", supplierTelephone)

'customerVatNumber = "0440058217"
documentTemplate = Replace(documentTemplate, "{customerVatNumber}", customerVatNumber)
documentTemplate = Replace(documentTemplate, "{customerTaxScheme}", customerTaxScheme)
'customerName = "Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv"
documentTemplate = Replace(documentTemplate, "{customerRegistrationName}", customerName)
'customerStreetName = "Grote Baan 141"
documentTemplate = Replace(documentTemplate, "{customerStreetName}", customerStreetName)
'customerCityName = "HERDERSEM"
documentTemplate = Replace(documentTemplate, "{customerCityName}", customerCityName)
'customerPostalZone = "9310"
documentTemplate = Replace(documentTemplate, "{customerPostalZone}", customerPostalZone)
'customerCountryCode = "BE"
documentTemplate = Replace(documentTemplate, "{customerCountryCode}", customerCountryCode)
'customerCountryCode = "BE"
documentTemplate = Replace(documentTemplate, "{customerSchemeId}", customerSchemeId)
'customerCompanyId = "04400580217"
documentTemplate = Replace(documentTemplate, "{customerCompanyId}", customerCompanyId)
'customerCompanyIdExtended = "BE04400580217"
documentTemplate = Replace(documentTemplate, "{customerCompanyIdExtended}", customerCompanyIdExtended)

'customerRegistrationName = "Zakenkantoor Hedwig Roelandt en Jos Vermoesen bv"
documentTemplate = Replace(documentTemplate, "{customerRegistrationName}", customerRegistrationName)
'customerElectronicMail = "info@rv.be"
documentTemplate = Replace(documentTemplate, "{customerElectronicMail}", customerElectronicMail)

paymentMeansCode = "1"
documentTemplate = Replace(documentTemplate, "{paymentMeansCode}", paymentMeansCode)
'paymentID = "024010022505"
documentTemplate = Replace(documentTemplate, "{paymentId}", paymentID)
'supplierIBAN = "BE50001683378618"
documentTemplate = Replace(documentTemplate, "{supplierIBAN}", supplierIBAN)
'supplierBIC = "GEBABEBB"
documentTemplate = Replace(documentTemplate, "{supplierBIC}", supplierBIC)
'paymentTerms = "024010022505"
documentTemplate = Replace(documentTemplate, "{paymentTerms}", paymentTerms)


documentTemplate = Replace(documentTemplate, "{taxTotalAmount}", taxGlobalTotalAmount)
documentTemplate = Replace(documentTemplate, "{taxableAmount}", taxableGlobalTotalAmount)
documentTemplate = Replace(documentTemplate, "{taxExclusiveAmount}", taxableGlobalTotalAmount)

globalTotalAmount = Trim(Dec((Val(taxableGlobalTotalAmount) + Val(taxGlobalTotalAmount)), MASK_EUR))
documentTemplate = Replace(documentTemplate, "{taxInclusiveAmount}", globalTotalAmount)
documentTemplate = Replace(documentTemplate, "{payableAmount}", globalTotalAmount)

Ktrl = ScrLeesBestandAlleTekst(invoiceTaxLineTemplate, PROGRAM_LOCATION + "xml-templates\ubl_be_3_0-invoicetaxline.xml")
If Ktrl = 0 Then
    MsgBox "Onverwachte situatie", vbCritical
End If

'0%, 6%, 12%, 21%
For T = 0 To 3
    If BTWEuroBasis(T) <> 0 Then
        taxSubtotalLine = invoiceTaxLineTemplate
        subTaxableAmount = Trim(Dec((BTWEuroBasis(T)), MASK_EUR))
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxableAmount}", subTaxableAmount)
                
        If T = 0 Then
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxAmount}", "0")
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxIndex}", "Z")
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxCategory}", "00")
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxPercent}", "0")
        Else
            subTaxAmount = Trim(Dec((BTWEuroBedrag(T)), MASK_EUR))
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxAmount}", subTaxAmount)
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxIndex}", "S")
            subTaxCategory = Trim(Dec((T), "00"))
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxCategory}", subTaxCategory)
            subTaxPercent = Mid(fmarBoxText("002", "2", Trim(Str(T))), 4, 4)
            taxSubtotalLine = Replace(taxSubtotalLine, "{taxPercent}", subTaxPercent)
        End If
        
        subTaxExclusiveAmount = Trim(Dec((BTWEuroBasis(T)), sy2))
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxExclusiveAmount}", subTaxExclusiveAmount)
        subTaxInclusiveAmount = Trim(Dec(BTWEuroBasis(T) + BTWEuroBedrag(T), MASK_EUR))
        taxSubtotalLine = Replace(taxSubtotalLine, "{taxInclusiveAmount}", subTaxInclusiveAmount)
        subPayableAmount = taxInclusiveAmount
        taxSubtotalLine = Replace(taxSubtotalLine, "{payableAmount}", subPayableAmount)
        taxSubtotalLine = Replace(taxSubtotalLine, "{currency}", invoiceCurrency)
        taxSubtotalLine = Replace(taxSubtotalLine, "{supplierTaxScheme}", supplierTaxScheme)
        
        listTaxSubtotalLines = listTaxSubtotalLines + taxSubtotalLine
        If T = 3 Then
        Else
            listTaxSubtotalLines = listTaxSubtotalLines + vbCrLf
        End If
    End If
Next
documentTemplate = Replace(documentTemplate, "<Vsoft>{taxsubtotals-template}</Vsoft>", listTaxSubtotalLines)

If Left(invoiceNumber, 2) = "V0" Then
    Ktrl = ScrLeesBestandAlleTekst(invoiceLineTemplate, PROGRAM_LOCATION + "xml-templates\ubl_be_3_0-invoiceline.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
Else
    Ktrl = ScrLeesBestandAlleTekst(invoiceLineTemplate, PROGRAM_LOCATION + "xml-templates\ubl_be_3_0-creditnoteline.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
End If

ublTT = 0
For T = 0 To VerkoopDetail.ListCount - 1
    If T > 0 Then
        listInvoiceLines = listInvoiceLines + vbCrLf
    End If
    thisInvoiceLine = invoiceLineTemplate
    GridText = VerkoopDetail.List(T)
    If Right(GridText, 1) = "2" Then
        If T = VerkoopDetail.ListCount - 1 Then
            'Stop
        End If
    Else
        ublTT = ublTT + 1
        lineCounter = Trim(Str(ublTT))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineCounter+1}", lineCounter)
        lineDescription = Trim(Mid(GridText, 1, 40))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineDescription}", lineDescription)

        GridText = Right(GridText, Len(GridText) - 41)
        dVeldUblInfo(4) = Val(Left(GridText, 11))          'price exclusive without discount
        GridText = Right(GridText, Len(GridText) - 12)
        dVeldUblInfo(6) = Val(Left(GridText, 7))           'quantity
        GridText = Right(GridText, Len(GridText) - 8)
        dVeldUblInfo(7) = Val(Left(GridText, 12))          'total exclusive with discount
        GridText = Right(GridText, Len(GridText) - 13)
        dVeldUblInfo(2) = Val(Left(GridText, 6))           'maatstaf anders default 1.0
        GridText = Right(GridText, Len(GridText) - 7)
        VeldUblInfo(8) = Left(GridText, 1)                 '9 ??
        GridText = Right(GridText, Len(GridText) - 2)
        dVeldUblInfo(5) = Val(Left(GridText, 3))           'korting
        GridText = Right(GridText, Len(GridText) - 4)
        VeldUblInfo(9) = Left(GridText, 1)                 'vat id
        GridText = Right(GridText, Len(GridText) - 2)
        'If VeldUblInfo(9) = "6" Then
        '    VeldUblInfo(9) = "0"
        'End If
        dVeldUblInfo(3) = Val(Left(GridText, 7))           'number ledgeraccount
        GridText = Right(GridText, Len(GridText) - 8)
        VeldUblInfo(0) = Left(GridText, 13)                'product reference id or description
                        
        lineQuantity = Trim(Dec(dVeldUblInfo(6), MASK_EUR))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineQuantity}", lineQuantity)
                
        'tmpString = "701000;701000;701000"
        'lineAccountingcost = Split(tmpString, ";")
        'tmpString = "8.4000;83.8951;343.5600"
        'lineTaxAmountAccurate = Split(tmpString, ";")
        'tmpString = "1-701000;1-701000;1-701000"
        'lineSellerItemIdentification = Split(tmpString, ";")
        'tmpString = "03;03;03"
        'lineTaxCategory = Split(tmpString, ";")
        'tmpString = "21.00;21.00;21.00"
        'lineTaxPercentage = Split(tmpString, ";")

        lineAccountingcost = Trim(Str(dVeldUblInfo(3)))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineAccountingCost}", lineAccountingcost)
        
        If VeldUblInfo(9) = "6" Then
            VeldUblInfo(9) = "0"
            
            linePriceAmount = Trim(Dec((dVeldUblInfo(7) / dVeldUblInfo(6)), MASK_EUR)) + "00"
            thisInvoiceLine = Replace(thisInvoiceLine, "{linePriceAmount}", linePriceAmount)
                    
            lineExtentionAmount = Trim(Dec(dVeldUblInfo(7), MASK_EUR)) + "00"
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineExtensionAmount}", lineExtentionAmount)
        
            thisInvoiceLine = Replace(thisInvoiceLine, "{taxIndex}", "Z")
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxCategory}", "00")
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxPercentage}", "0")
            
            lineTaxAmountAccurate = "0.0000"
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxAmountAccurate}", lineTaxAmountAccurate)
            
        Else
            linePriceAmount = Trim(Dec((dVeldUblInfo(7) / dVeldUblInfo(6)), MASK_EUR + "00"))
            thisInvoiceLine = Replace(thisInvoiceLine, "{linePriceAmount}", linePriceAmount)
        
            lineExtentionAmount = Trim(Dec(dVeldUblInfo(7), MASK_EUR + "00"))
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineExtensionAmount}", lineExtentionAmount)
        
            thisInvoiceLine = Replace(thisInvoiceLine, "{taxIndex}", "S")
            lineTaxCategory = Dec(Val(VeldUblInfo(9)), "00")
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxCategory}", lineTaxCategory)
        
            lineTaxPercentage = Mid(fmarBoxText("002", "2", VeldUblInfo(9)), 4, 4)
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxPercentage}", lineTaxPercentage)
                
            lineTaxAmountAccurate = Trim(Dec((dVeldUblInfo(7) * Val(lineTaxPercentage) / 100), MASK_EUR + "00"))
            thisInvoiceLine = Replace(thisInvoiceLine, "{lineTaxAmountAccurate}", lineTaxAmountAccurate)
        End If
        
        lineSellerItemIdentification = Trim(VeldUblInfo(0))
        thisInvoiceLine = Replace(thisInvoiceLine, "{lineSellerItemIdentification}", lineSellerItemIdentification)
        
        thisInvoiceLine = Replace(thisInvoiceLine, "{currency}", invoiceCurrency)
        thisInvoiceLine = Replace(thisInvoiceLine, "{supplierTaxScheme}", supplierTaxScheme)
    
        listInvoiceLines = listInvoiceLines + thisInvoiceLine
    End If
Next

If Left(invoiceNumber, 2) = "V0" Then
    documentTemplate = Replace(documentTemplate, "<Vsoft>{invoicelines-template}</Vsoft>", listInvoiceLines)
Else
    documentTemplate = Replace(documentTemplate, "<Vsoft>{creditnotelines-template}</Vsoft>", listInvoiceLines)
End If

Ktrl = ScrMaakTekstBestand(documentTemplate, LOCATION_COMPANYDATA + "peppol\out\" + invoiceNumber + ".xml")
If Ktrl = 0 Then
    MsgBox "Fout bij bewaren van " + invoiceNumber + ".xml", vbCritical
Else
    XmlInvoiceGenerateUBLBE30 = True
End If

End Function


Function CheckDocument(sellerDoc As String) As String

    CheckDocument = ""
            
    Set rsSellerUBL = New ADODB.Recordset
    On Error Resume Next
    Err = 0
    rsSellerUBL.CursorLocation = adUseClient
                
    Dim sSQL As String
        
    sSQL = "SELECT "
    sSQL = sSQL + "Klanten.A110, Klanten.A100, Klanten.v404, Klanten.A161, Klanten.v150, "
    sSQL = sSQL + "Dokumenten.v033, Dokumenten.v035, Dokumenten.v036, Dokumenten.v037, Dokumenten.v249, Dokumenten.v405 "
    sSQL = sSQL + "FROM  Klanten, Dokumenten "
    sSQL = sSQL + "WHERE Dokumenten.v034 = 'K'+Klanten.A110 "
    sSQL = sSQL + "AND Dokumenten.v033 = '" + Mid(sellerDoc, 1, 11) + "' "
       
    Screen.MousePointer = vbHourglass
    rsSellerUBL.Open sSQL, adntDB, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        Exit Function
    ElseIf rsSellerUBL.RecordCount = 0 Then
        CheckDocument = "Niet gevonden"
        Exit Function
    ElseIf rsSellerUBL.RecordCount = 1 Then
        'check status
        Dim sentID As String
        sentID = Trim(rsSellerUBL("v405"))
        If sentID = "" Then
            CheckDocument = "Te verzenden"
        Else
            CheckDocument = "Verzonden   "
            Dim totalPaid As Double
            Dim toPay As Double
                        
            totalPaid = Val(rsSellerUBL("v037"))
            toPay = Val(rsSellerUBL("v249"))
                        
            If (toPay - totalPaid) = 0 Then
                CheckDocument = CheckDocument & "Betaald"
                Exit Function
            Else
                CheckDocument = CheckDocument & "Openstaand"
                Exit Function
            End If
        End If
    Else
        If rsSellerUBL.RecordCount > 1 Then
            CheckDocument = "whatsgoingon?"
            MsgBox "Meerdere documenten gevonden, onlogische situatie.", vbExclamation
            Exit Function
        End If
    End If
    
End Function

Function CheckCustomerDocuments(customerCode As String) As Boolean
       
    Dim valueInRecordV407 As String
    Dim checkWithVatNumber As String
    
    CheckCustomerDocuments = False
    
    Dim rsAny As ADODB.Recordset
    Set rsAny = New ADODB.Recordset
    
    On Error Resume Next
    Err = 0
    rsAny.CursorLocation = adUseClient
    
    Msg = "SELECT A110, v150, A161, v404, v407, A100 FROM Klanten WHERE A110 = '" & customerCode & "'"
    SnelHelpPrint Msg, BL_LOGGING
    Screen.MousePointer = vbHourglass
    rsAny.Open Msg, adntDB, adOpenDynamic, adLockOptimistic
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    ElseIf rsAny.RecordCount = 1 Then
        If IsNull(rsAny("v407")) Or rsAny("v407").Value = "" Then
            'First check with 0208:"
            checkWithVatNumber = "0208:" + rsAny("V404")
            valueInRecordV407 = CheckPeppolRegistration(checkWithVatNumber)
            'MsgBox "lengte: " + Str(Len(valueInRecordV410))
            'TODO: definitly to check for Peppol readiness
            If Len(valueInRecordV407) < 300 Then
                Msg = "Gecontroleerd met code: "
                Msg = Msg & checkWithVatNumber & vbCrLf & vbCrLf
                Msg = Msg & "Mogelijk geen Peppol Registratie" & vbCrLf & vbCrLf
                Msg = Msg & "Tot slot controleren met verouderde 9925:BE"
                MsgBox Msg, vbInformation
                checkWithVatNumber = "9925:BE" + rsAny("V404")
                valueInRecordV407 = CheckPeppolRegistration(checkWithVatNumber)
                'MsgBox "lengte: " + Str(Len(valueInRecordV410))
                'TODO: definitly to check for Peppol readiness
                If Len(valueInRecordV407) < 300 Then
                    CheckCustomerDocuments = False
                Else
                    CheckCustomerDocuments = True
                End If
            Else
                CheckCustomerDocuments = True
            End If
            rsAny("v407") = valueInRecordV407
            rsAny.Update
        ElseIf Len(rsAny("v407")) > 500 And Me.CheckBoxAlwaysPeppolRefresh.Value = vbUnchecked Then
            CheckCustomerDocuments = True
        ElseIf Me.CheckBoxAlwaysPeppolRefresh.Value = vbChecked Then 'always refresh!
            'First check with 0208:"
            checkWithVatNumber = "0208:" + rsAny("V404")
            valueInRecordV407 = CheckPeppolRegistration(checkWithVatNumber)
            'MsgBox "lengte: " + Str(Len(valueInRecordV410))
            'TODO: definitly to check for Peppol readiness
            If Len(valueInRecordV407) < 300 Then
                Msg = "Gecontroleerd met code: "
                Msg = Msg & checkWithVatNumber & vbCrLf & vbCrLf
                Msg = Msg & "Mogelijk geen Peppol Registratie" & vbCrLf & vbCrLf
                Msg = Msg & "Tot slot controleren met verouderde 9925:BE"
                MsgBox Msg, vbInformation
                checkWithVatNumber = "9925:BE" + rsAny("V404")
                valueInRecordV407 = CheckPeppolRegistration(checkWithVatNumber)
                'MsgBox "lengte: " + Str(Len(valueInRecordV410))
                'TODO: definitly to check for Peppol readiness
                If Len(valueInRecordV407) < 300 Then
                    CheckCustomerDocuments = False
                Else
                    CheckCustomerDocuments = True
                End If
            Else
                CheckCustomerDocuments = True
            End If
            rsAny("v407") = valueInRecordV407
            rsAny.Update
        End If
        rsAny.Close
    End If
    Set rsAny = Nothing
    
End Function

