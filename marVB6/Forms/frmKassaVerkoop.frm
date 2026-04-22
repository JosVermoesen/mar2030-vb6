VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmKassaVerkoop 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kassaverkoop"
   ClientHeight    =   5715
   ClientLeft      =   435
   ClientTop       =   1365
   ClientWidth     =   8205
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5715
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "&Inbreng"
      TabPicture(0)   =   "frmKassaVerkoop.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(20)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(12)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblInBtw"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblExBtw"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LblIn2Btw"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LblEx2Btw"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LblEurCash"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "LbCash"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdWissel"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cbEpsonTM"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdFactuurListing"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdSnelRekenen"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdDagAfsluiting"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cbVandaag"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdLijst"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "CmbAfdruk"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "CmdStock"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Annuleren"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Afsluiten"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Schoonvegen"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdKassaKiezen"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "KassaDetail"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdSwitch"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "&Ticket"
      TabPicture(1)   =   "frmKassaVerkoop.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblVierkant"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "txtTicket"
      Tab(1).Control(3)=   "cmdPrintAf"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "&Betalen"
      TabPicture(2)   =   "frmKassaVerkoop.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fBetaalwijze"
      Tab(2).Control(1)=   "txtTicketTotaal"
      Tab(2).Control(2)=   "cbVerwerken2"
      Tab(2).Control(3)=   "fContant"
      Tab(2).ControlCount=   4
      Begin VB.Frame fContant 
         Caption         =   "Contant"
         Height          =   1455
         Left            =   -74760
         TabIndex        =   40
         Top             =   3120
         Width           =   5715
         Begin VB.CommandButton cmdWissel2 
            Caption         =   "Wissel Hulp"
            Enabled         =   0   'False
            Height          =   795
            Left            =   120
            TabIndex        =   45
            Top             =   480
            Width           =   1035
         End
         Begin VB.TextBox txtOntvangenBEF 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   345
            Left            =   1920
            TabIndex        =   44
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtOntvangenEUR 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   1920
            TabIndex        =   43
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox txtTeruggaveBEF 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   345
            Left            =   4200
            TabIndex        =   42
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtTeruggaveEUR 
            Alignment       =   1  'Right Justify
            Height          =   345
            Left            =   4200
            TabIndex        =   41
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ontvangen"
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   51
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Teruggave"
            Height          =   315
            Index           =   1
            Left            =   3600
            TabIndex        =   50
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "BEF"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   1320
            TabIndex        =   49
            Top             =   960
            Width           =   555
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "EUR"
            Height          =   315
            Index           =   3
            Left            =   1320
            TabIndex        =   48
            Top             =   600
            Width           =   555
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "BEF (info)"
            Enabled         =   0   'False
            Height          =   315
            Index           =   4
            Left            =   3240
            TabIndex        =   47
            Top             =   960
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "EUR"
            Height          =   315
            Index           =   5
            Left            =   3660
            TabIndex        =   46
            Top             =   600
            Width           =   495
         End
      End
      Begin VB.CommandButton cbVerwerken2 
         Caption         =   "Verwerken"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -71460
         TabIndex        =   39
         Top             =   2220
         Width           =   2415
      End
      Begin VB.TextBox txtTicketTotaal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   -71460
         TabIndex        =   38
         Top             =   600
         Width           =   2475
      End
      Begin VB.Frame fBetaalwijze 
         Caption         =   "Betaalwijze"
         Height          =   2535
         Left            =   -74760
         TabIndex        =   29
         Top             =   480
         Width           =   3195
         Begin VB.OptionButton obBetaling 
            Caption         =   "BEF-Betaling contant :"
            Enabled         =   0   'False
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
            Index           =   5
            Left            =   180
            TabIndex        =   35
            Top             =   2160
            Width           =   2715
         End
         Begin VB.OptionButton obBetaling 
            Caption         =   "CreditCard           :"
            Enabled         =   0   'False
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
            Index           =   3
            Left            =   180
            TabIndex        =   34
            Top             =   1260
            Width           =   2715
         End
         Begin VB.OptionButton obBetaling 
            Caption         =   "Cheque               :"
            Enabled         =   0   'False
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
            Index           =   4
            Left            =   180
            TabIndex        =   33
            Top             =   1560
            Width           =   2715
         End
         Begin VB.OptionButton obBetaling 
            Caption         =   "Bancontact Maestro   :"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   180
            TabIndex        =   32
            Top             =   900
            Width           =   2655
         End
         Begin VB.OptionButton obBetaling 
            Caption         =   "Proton               :"
            Enabled         =   0   'False
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
            Index           =   1
            Left            =   180
            TabIndex        =   31
            Top             =   645
            Width           =   2655
         End
         Begin VB.OptionButton obBetaling 
            Caption         =   "EUR-Betaling contant :"
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
            Index           =   0
            Left            =   180
            TabIndex        =   30
            Top             =   360
            Value           =   -1  'True
            Width           =   2655
         End
      End
      Begin VB.CommandButton cmdPrintAf 
         Caption         =   "Afdrukken"
         Height          =   375
         Left            =   -68340
         TabIndex        =   28
         Top             =   4680
         Width           =   1395
      End
      Begin VB.TextBox txtTicket 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4635
         Left            =   -73620
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Text            =   "frmKassaVerkoop.frx":0054
         Top             =   480
         Width           =   5175
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
         Height          =   825
         Left            =   7200
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   4320
         Width           =   885
      End
      Begin VB.ListBox KassaDetail 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   15
         Top             =   630
         Width           =   7995
      End
      Begin VB.CommandButton cmdKassaKiezen 
         Caption         =   "Kassa"
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
         Left            =   120
         TabIndex        =   0
         Top             =   3330
         Width           =   1245
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
         Height          =   375
         Left            =   5280
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   4800
         Width           =   885
      End
      Begin VB.CommandButton Afsluiten 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Verwerken"
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
         Left            =   120
         TabIndex        =   13
         Top             =   4290
         Width           =   1245
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
         Height          =   375
         Left            =   6240
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   4800
         Width           =   915
      End
      Begin VB.CommandButton CmdStock 
         Caption         =   "&Stock artikel"
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
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3810
         Width           =   1245
      End
      Begin VB.CommandButton CmbAfdruk 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Extra &Afdruk"
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
         Left            =   120
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4770
         Width           =   1245
      End
      Begin VB.CommandButton cmdLijst 
         Caption         =   "&Lijst"
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
         Left            =   3000
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   4290
         Width           =   795
      End
      Begin VB.CheckBox cbVandaag 
         Caption         =   "&Enkel Vandaag"
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
         Left            =   1440
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   4770
         Value           =   1  'Checked
         Width           =   1755
      End
      Begin VB.CommandButton cmdDagAfsluiting 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Dag Afsluiting !"
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
         Height          =   855
         Left            =   1440
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3330
         Width           =   2355
      End
      Begin VB.CommandButton cmdSnelRekenen 
         Caption         =   "SnelRekenen"
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   4320
         Width           =   1875
      End
      Begin VB.CommandButton cmdFactuurListing 
         Caption         =   "Factuur Listing"
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
         Left            =   1440
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   4290
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.CheckBox cbEpsonTM 
         Caption         =   "Epson TM spoolbestand "
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
         Left            =   1440
         TabIndex        =   4
         Top             =   5010
         Width           =   2175
      End
      Begin VB.CommandButton cmdWissel 
         Caption         =   "WisselHulp (buiten kas)"
         Enabled         =   0   'False
         Height          =   855
         Left            =   4080
         TabIndex        =   3
         Top             =   4320
         Width           =   1155
      End
      Begin VB.Label LbCash 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EUR incl.Btw Cash"
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
         Left            =   3840
         TabIndex        =   53
         Top             =   4020
         Width           =   1485
      End
      Begin VB.Label LblEurCash 
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
         Height          =   345
         Left            =   5400
         TabIndex        =   52
         ToolTipText     =   "Totaal dokument in EURO inclusief BTW"
         Top             =   3960
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Bef<>Eur"
         Height          =   255
         Left            =   -67980
         TabIndex        =   37
         Top             =   540
         Width           =   915
      End
      Begin VB.Label lblVierkant 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   -68040
         TabIndex        =   36
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Prijs"
         Height          =   240
         Index           =   0
         Left            =   5160
         TabIndex        =   25
         Top             =   420
         Width           =   465
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Aantal"
         Height          =   240
         Index           =   4
         Left            =   5880
         TabIndex        =   24
         Top             =   420
         Width           =   630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LijnTotaal"
         Height          =   240
         Index           =   2
         Left            =   6870
         TabIndex        =   23
         Top             =   420
         Width           =   1110
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "&Detail lijnen"
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
         Left            =   240
         TabIndex        =   22
         Top             =   360
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
         Height          =   345
         Left            =   6780
         TabIndex        =   21
         ToolTipText     =   "Totaal dokument in EURO exclusief BTW"
         Top             =   3630
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
         Height          =   345
         Left            =   5400
         TabIndex        =   20
         ToolTipText     =   "Totaal dokument in EURO inclusief BTW"
         Top             =   3630
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
         Height          =   345
         Left            =   6780
         TabIndex        =   19
         ToolTipText     =   "Totaal dokument in BEF exclusief BTW"
         Top             =   3270
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
         Height          =   345
         Left            =   5400
         TabIndex        =   18
         ToolTipText     =   "Totaal dokument in BEF inclusief BTW"
         Top             =   3270
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
         Left            =   3870
         TabIndex        =   17
         Top             =   3690
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
         Left            =   3870
         TabIndex        =   16
         Top             =   3330
         Width           =   1485
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   5325
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "7:13"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "21/04/2026"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "0000"
            TextSave        =   "0000"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmKassaVerkoop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim dokumentType         As String * 2
Dim dokumentHistoriek As String * 11

Dim KassaDetailTitel(8) As String
Dim dokumentSleutel As String * 11
Dim KlantRekening As String * 7

Dim VerkoopFLG As Integer
Dim AfdrukFlag As Integer
Dim AantalEx As String * 2
Dim DefaultVerkoop As String * 7
Dim KontaktPersoon As Integer
Dim rbtwVAK(10) As String * 7
Dim Vr As Integer
Dim TempoVerkoopGridText As String

Dim VKKTBtwBEF As Double
Dim VKKTGoedBEF As Double
Dim VKKTBtwEUR As Double
Dim VKKTGoedEUR As Double

Dim BTWEuroBasis(3) As Double
Dim BTWEuroBedrag(3) As Double

Dim BTWBasis(3) As Double
Dim BTWBedrag(3) As Double
Dim TotaalBTW As Double
Dim TotaalUitvoer As Double

Dim dMuntK              As Double
Dim sMunt               As String * 3
Dim sMuntKlant          As String * 3

Dim VsoftVanaf          As String
Dim VsoftTot            As String
Dim BeginXbox           As Integer
Dim BeginYbox           As Integer
Dim OVSStrooklijnen     As Long

Dim ForFait             As Integer
Dim SteedsDrukken       As Boolean

Dim rsDetail            As New ADODB.Recordset
Dim cnnKassa            As New ADODB.Connection

Dim KassaPrinter        As Printer
Dim KassaNaam           As String


Private Function WegSchrijfFout() As Boolean

Dim T               As Integer
Dim AantalStuks     As Single
Dim DetailMEMO      As Variant

WegSchrijfFout = True
DetailMEMO = ""
For T = 0 To KassaDetail.ListCount - 1
    'KassaDetail.ListIndex = T
    DetailMEMO = DetailMEMO & KassaDetail.List(T) & vbCr
    Select Case Right(KassaDetail.List(T), 1)
        Case "0" 'Produkt !!
            '1, 40      Omschrijving
            '42, 11     BedragGoed
            '54, 7      Aantal
            '62, 12     Totaal
            '75, 6      Verpakking
            '90, 7      VerkoopRekening
            '98, 13     ProduktNummer
            
            bGet TABLE_PRODUCTS, 0, Mid(KassaDetail.List(T), 98, 13)
            If Ktrl Then
                Beep
                MsgBox "stop"
            Else
                RecordToVeld TABLE_PRODUCTS
                AantalStuks = Val(Mid(KassaDetail.List(T), 54, 7)) * Val(Mid(KassaDetail.List(T), 75, 6))
                vBib TABLE_PRODUCTS, Dec(AantalStuks + Val(vBibTekst(TABLE_PRODUCTS, "#v120 #")), MASK_SY(2)), "v120"
                If bhEuro Then
                    vBib TABLE_PRODUCTS, Dec((AantalStuks * Val(vBibTekst(TABLE_PRODUCTS, "#e113 #"))) + Val(vBibTekst(TABLE_PRODUCTS, "#e122 #")), MASK_SY(2)), "e122"
                Else
                    vBib TABLE_PRODUCTS, Dec((AantalStuks * Val(vBibTekst(TABLE_PRODUCTS, "#v113 #"))) + Val(vBibTekst(TABLE_PRODUCTS, "#v122 #")), MASK_SY(2)), "v122"
                End If
                bUpdate TABLE_PRODUCTS, 0
                If Ktrl Then
                    GoTo FoutMelding
                End If
            End If
        Case Else
            MsgBox "stop"
    End Select
Next
rsDetail.AddNew
rsDetail("AantalProducten") = KassaDetail.ListCount
rsDetail("EuroExclusief") = CDbl(LblEx2Btw)
rsDetail("BefExclusief") = CDbl(LblExBtw)
rsDetail("EuroBTW") = CDbl(LblIn2Btw) - CDbl(LblEx2Btw)
rsDetail("BefBTW") = CDbl(LblInBtw) - CDbl(LblExBtw)
rsDetail("Dag") = DATE_KEY(Format(Now, "dd/mm/yyyy"))
rsDetail("DagUur") = Now
rsDetail("Info") = DetailMEMO
For COUNT_TO = 0 To 5
    If obBetaling(COUNT_TO).Value = True Then
        rsDetail("BetaalType") = Trim$(Str$(COUNT_TO)) & ":" & obBetaling(COUNT_TO).Caption
        Exit For
    End If
Next
rsDetail("BetaaldEUR") = Val(txtOntvangenEUR.text)
rsDetail("BetaaldBEF") = Val(txtOntvangenBEF.text)
rsDetail("TerugEUR") = Val(txtTeruggaveEUR.text)

rsDetail.Update
WegSchrijfFout = False
Exit Function

FoutMelding:
MsgBox "Stopkode " + Str$(Ktrl) + " tijdens het bijwerken." + vbCrLf + "Hierna worden alle bewerkingen genegeerd." + vbCrLf + vbCrLf + "Begin opnieuw of neem kontakt op 053/21.59.25.", vbCritical

End Function

Private Sub Afsluiten_Click()

Err = 0
If UCase(String99(READING, 130)) = "OFF" Then
Else
    If cbEpsonTM.Value Then
        X = PrintAf(False, False)
    Else
        X = PrintAf(True, False)
    End If
End If

If bhEuro Then
    If cmdSwitch.Caption = "Ingave in BEF" Then
        cmdSwitch_Click
    End If
Else
    If cmdSwitch.Caption = "Ingave in EUR" Then
        cmdSwitch_Click
    End If
End If
Err = 0
Msg = "Stock bijwerken (volgende klant)" + vbCr + vbCr
Msg = Msg + "Bent U zeker."
KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
If KtrlBox = vbNo Then Exit Sub

If BoekingKassa = False Then
    MsgBox "Verwerking zonder succes"
End If

End Sub
Private Sub Annuleren_Click()

If KassaDetail.ListCount Then
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



Private Sub cbEpsonTM_Click()

BeWaarTekst "KassaVerkoop", "EpsonTMModus", cbEpsonTM.Value
If cbEpsonTM.Value = 1 Then
    If Dir("C:\Program Files\Epson\VbSample.exe") = "" Then
        MsgBox "Printspooler is nog niet geïnstalleerd", vbCritical
        cbEpsonTM.Value = 0
    End If
End If

End Sub

Private Sub cbVerwerken2_Click()

    Err = 0
    Msg = "Verkoop afsluiten met betalingen:" & vbCr & vbCr
    For COUNT_TO = 0 To 5
        If obBetaling(COUNT_TO).Value = True Then
            Msg = Msg & Str$(COUNT_TO) & ":" & obBetaling(COUNT_TO).Caption & vbCr & vbCr
            Exit For
        End If
    Next
    If txtTeruggaveEUR.Visible = True Then
        If Val(txtOntvangenBEF.text) <> 0 Then
            Msg = Msg & "Ontvangen BEF " & txtOntvangenBEF.text & vbCr
        End If
        If Val(txtOntvangenEUR.text) <> 0 Then
            Msg = Msg & "Ontvangen EUR " & txtOntvangenEUR.text & vbCr
        End If
        Msg = Msg & "Teruggave EUR " & txtTeruggaveEUR.text & vbCr & vbCr
    End If
    Msg = Msg + "Bent U zeker."
    KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
    If KtrlBox = vbNo Then Exit Sub
    
    Err = 0
    If UCase(String99(READING, 130)) = "OFF" Then
    Else
        If cbEpsonTM.Value Then
            X = PrintAf(False, False)
        Else
            X = PrintAf(True, False)
        End If
    End If

    If bhEuro Then
        If cmdSwitch.Caption = "Ingave in BEF" Then
            cmdSwitch_Click
        End If
    Else
        If cmdSwitch.Caption = "Ingave in EUR" Then
            cmdSwitch_Click
        End If
    End If
    
    If BoekingKassa = False Then
        MsgBox "Verwerking zonder succes"
    End If
    
End Sub

Private Sub CmbAfdruk_Click()

If cbEpsonTM.Value Then
    X = PrintAf(False, False)
Else
    X = PrintAf(True, False)
End If

End Sub

Function PrintAf(naarPrinter As Boolean, inScherm As Boolean) As Boolean

Dim BTWBasisHier        As Currency
Dim BTWHier             As Currency
Dim Totaal              As Currency

Dim PPStuk              As Double
Dim PPPakket            As Double
Dim BTWPStuk            As Double

Dim AantalStuks         As Single

Dim dVeldInfo(10)       As Double
Dim Omschrijving        As String * 40
Dim ArtikelCode         As String * 13
Dim BTWKode             As String * 1
Dim BtwTekst            As String * 4

Dim Maatstaf            As String * 5

Dim MaxVoet             As Integer
Dim TabJe               As Integer

Dim aa                  As String
Dim LijnString As String
Dim TicketVar As Variant
Dim flHier As Integer
Dim TotaalBetaaldEUR As Currency

If naarPrinter Then
    Set Printer = Printers(KassaPrinterNr)
    On Error Resume Next
    Printer.PaperBin = LaadTekst(App.Title, "KASSAPRINTER")
    'TabJe = Val(String99(READING, 131))
    Printer.Print vbCrLf;
End If
TicketVar = vbCrLf

LijnString = Format(Now, "dd/mm/yyyy hh:mm:ss")
LijnString = LijnString + " " + StatusBar1.Panels(3).text
LijnString = LijnString + " " + Trim$(KassaNaam)
If naarPrinter Then
    Printer.Print LijnString;
    Printer.Print vbCrLf; vbCrLf;
End If
TicketVar = TicketVar & LijnString & vbCrLf & vbCrLf


LijnString = "ARTICEL CODE  VERPAKKING  AANTAL   PRIJS"
If naarPrinter Then
    Printer.Print LijnString; vbCrLf;
End If
TicketVar = TicketVar & LijnString & vbCrLf

LijnString = "OMSCHRIJVING"
If naarPrinter Then
    Printer.Print LijnString; vbCrLf;
    Printer.Print String(40, "-"); vbCrLf; vbCrLf;
End If
TicketVar = TicketVar & LijnString & vbCrLf & String(40, "-") & vbCrLf & vbCrLf

For pr = 0 To KassaDetail.ListCount - 1
    aa = KassaDetail.List(pr)
    GoSub FilterVelden
    LijnString = ArtikelCode + Dec((dVeldInfo(2)), "###0.0") + " " + UCase(Maatstaf)
        BTWBasisHier = dVeldInfo(7)
        BTWHier = BTWBasisHier * Val(BtwTekst) / 100
        AantalStuks = dVeldInfo(6)
        PPStuk = dVeldInfo(4) + (dVeldInfo(4) * Val(BtwTekst) / 100)
        
        If InStr(cmdSwitch.Caption, "EUR") Then
            PPPakket = Round(BTWBasisHier, 2) + Round(BTWHier, 2)
        Else
            PPPakket = Round(BTWBasisHier) + Round(BTWHier)
        End If
        
    LijnString = LijnString + Dec((AantalStuks), Right(MASK_EUR, 5))
    LijnString = LijnString + Dec(PPPakket, MASK_EUR)
    If naarPrinter Then
        Printer.Print LijnString; vbCrLf;
    End If
    TicketVar = TicketVar & LijnString & vbCrLf
        
    LijnString = UCase(Trim$(Omschrijving))
    If naarPrinter Then
        Printer.Print LijnString; vbCrLf;
    End If
    TicketVar = TicketVar & LijnString & vbCrLf
           
    If InStr(cmdSwitch.Caption, "EUR") Then
        Totaal = Totaal + Round(BTWBasisHier, 2) + Round(BTWHier, 2)
    Else
        Totaal = Totaal + Round(BTWBasisHier) + Round(BTWHier)
    End If
Next

If InStr(cmdSwitch.Caption, "EUR") Then
    'EUR administratie
    'If Round(Totaal, 2) <> CDbl(LblIn2Btw) Then
    '    MsgBox "Vierkantskontrole geeft " & Totaal, vbExclamation, "Kontroleer"
    'End If
    lblVierkant = Format(Totaal - CDbl(LblIn2Btw), "#,##0.00")
Else
    'BEF administratie
    'If Round(Totaal) <> CDbl(LblInBtw) Then
    '    MsgBox "Vierkantskontrole geeft " & Totaal, vbExclamation, "Kontroleer"
    'End If
    lblVierkant = Format(Totaal - CDbl(LblInBtw), "#,##0.00")
End If
    
LijnString = String$(40, "-")
If naarPrinter Then
    Printer.Print LijnString; vbCrLf;
End If
TicketVar = TicketVar & LijnString & vbCrLf


'LijnString = "T O T A A L             E U R " + Dec(CDbl(LblIn2Btw), MASK_EUR)
LijnString = "T O T A A L             E U R " + Dec(CDbl(LblEurCash), MASK_EUR)

If naarPrinter Then
    Printer.Print LijnString;
    Printer.Print vbCrLf;
End If

'TicketVar = TicketVar & LijnString & vbCrLf
'LijnString = "                        B E F " + Dec(CDbl(LblInBtw), MASK_EUR)
'If naarPrinter Then
'    Printer.Print LijnString;
'    Printer.Print vbCrLf; vbCrLf;
'End If
TicketVar = TicketVar & LijnString & vbCrLf & vbCrLf
    
LijnString = ""
If SSTab1.Tab = 2 Or SSTab1.Tab = 1 Then
    For COUNT_TO = 0 To 5
        If obBetaling(COUNT_TO).Value = True Then
            TotaalBetaaldEUR = Val(txtOntvangenEUR.text) + Round(Val(txtOntvangenBEF.text) / EURO, 2)
            LijnString = obBetaling(COUNT_TO).Caption & "  E U R " & Dec((TotaalBetaaldEUR), MASK_EUR)
            If naarPrinter Then
                Printer.Print LijnString;
                Printer.Print vbCrLf;
            End If
            TicketVar = TicketVar & LijnString & vbCrLf
            If Val(txtTeruggaveEUR.text) <> 0 Then
                LijnString = "Teruggave               E U R " & Dec(Val(txtTeruggaveEUR.text), MASK_EUR)
                If naarPrinter Then
                    Printer.Print LijnString;
                    Printer.Print vbCrLf; vbCrLf;
                End If
                TicketVar = TicketVar & LijnString & vbCrLf & vbCrLf
            Else
                If naarPrinter Then
                    Printer.Print vbCrLf;
                End If
                TicketVar = TicketVar & vbCrLf
            End If
            Exit For
        End If
    Next
End If

MaxVoet = Val(String99(READING, 132))
If MaxVoet > 8 Then MaxVoet = 8
If MaxVoet = 0 Then
    If naarPrinter Then
        Printer.Print vbCrLf; vbCrLf;
    End If
    TicketVar = TicketVar & vbCrLf & vbCrLf
Else
    For T = 1 To MaxVoet
        LijnString = String99(READING, 136 + T)
        If naarPrinter Then
            Printer.Print LijnString; vbCrLf;
        End If
        TicketVar = TicketVar & LijnString & vbCrLf
    Next
    If naarPrinter Then
        Printer.Print vbCrLf; vbCrLf;
    End If
    TicketVar = TicketVar & vbCrLf & vbCrLf
End If
For T = 133 To 136
    LijnString = String99(READING, T)
    If naarPrinter Then
        Printer.Print LijnString; vbCrLf;
    End If
    TicketVar = TicketVar & LijnString & vbCrLf
Next
If naarPrinter Then
    Printer.EndDoc
Else
    If inScherm = True Then
        txtTicket.text = TicketVar
        Exit Function
    End If
    If Dir("c:\kassaticket.txt") <> "" Then
        Kill "c:\kassaticket.txt"
    End If
    GoSub DOSprintTicket
End If
Exit Function

DOSprintTicket:
flHier = FreeFile
Open "c:\kassaticket.txt" For Append As flHier
    Print #flHier, TicketVar;
Close flHier

Dim MyLongID As Long
On Error Resume Next
Err = 0
MyLongID = Shell("C:\Program Files\Epson\VbSample.exe", 1)
If Err Then
    MsgBox Error & vbCr & vbCr & "C:\Program Files\Epson\VbSample.exe"
Else
    DoEvents
    AppActivate MyLongID
    SendKeys "{ENTER}", True
    SendKeys "%{F4}", True
End If
Return

FilterVelden:
Omschrijving = Mid(aa, 1, 40)
aa = Right(aa, Len(aa) - 41)

'Prijs per verpakking
dVeldInfo(4) = Val(Left(aa, 11))
aa = Right(aa, Len(aa) - 12)

'Aantal in verpakking
dVeldInfo(6) = Val(Left(aa, 7))
aa = Right(aa, Len(aa) - 8)

'Totaal excl.
dVeldInfo(7) = Val(Left(aa, 12))
aa = Right(aa, Len(aa) - 13)

'Aantal verpakkingen
dVeldInfo(2) = Val(Left(aa, 6))
aa = Right(aa, Len(aa) - 7)

'
Maatstaf = Mid(fmarBoxText("004", "2", Left(aa, 1)), 4)

'VeldInfo(8) = Left(aa, 1)
aa = Right(aa, Len(aa) - 2)
dVeldInfo(5) = Val(Left(aa, 3))
aa = Right(aa, Len(aa) - 4)


'VeldInfo(9) = Left(aa, 1)
'BtwTekst=
Mid(BtwTekst, 1) = Mid(fmarBoxText("002", "2", Left(aa, 1)), 4, 4)

aa = Right(aa, Len(aa) - 2)
dVeldInfo(3) = Val(Left(aa, 7))
aa = Right(aa, Len(aa) - 8)
ArtikelCode = Left(aa, 13)
Return

End Function

Private Sub cmdDagAfsluiting_Click()

Dim DetailS As Variant

    If KassaDetail.ListCount <> 0 Then
        MsgBox "Er zijn nog lijnen te verwerken (schoonvegen indien geen verwerking nodig a.u.b.)", vbExclamation
        Exit Sub
    End If
    Msg = "Dagafsluiting" & vbCr & vbCr
    Msg = Msg & "Bent U zeker."
    KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
    If KtrlBox = vbNo Then Exit Sub
    
     If bhEuro Then
        If cmdSwitch.Caption = "Ingave in BEF" Then
            cmdSwitch_Click
        End If
    Else
        If cmdSwitch.Caption = "Ingave in EUR" Then
            cmdSwitch_Click
        End If
    End If
    
    If Not DATE_CHECK(Format(Now, "dd/mm/yyyy"), PERIODAS_TEXT) Then
        MsgBox "Kontroleer Dagdatum en actieve periode a.u.b. vooraleer verder te gaan", vbExclamation, Now & " ? " & BYPERDAT.PeriodeBoekjaar & " ?"
        Exit Sub
    End If
    dokumentSleutel = SleutelDok(11)
    VKKTBtwBEF = 0
    VKKTGoedBEF = 0
    VKKTBtwEUR = 0
    VKKTGoedEUR = 0
    rsDetail.MoveFirst
    Do While Not rsDetail.EOF
        DetailS = rsDetail("Info")
        Do While DetailS <> ""
            KassaDetail.AddItem Mid(DetailS, 1, InStr(DetailS, vbCr) - 1)
            DetailS = Mid(DetailS, InStr(DetailS, vbCr) + 1)
        Loop
        VKKTBtwBEF = VKKTBtwBEF + rsDetail("BefBTW")
        VKKTGoedBEF = VKKTGoedBEF + rsDetail("BefExclusief")
        VKKTBtwEUR = VKKTBtwEUR + rsDetail("EuroBtw")
        VKKTGoedEUR = VKKTGoedEUR + rsDetail("EuroExclusief")
        rsDetail.MoveNext
    Loop
    RefreshBTW
   
    Msg = "Hierna wordt de boekhouding bijgewerkt en kan U vandaag" & vbCr
    Msg = Msg & "niets meer inbrengen voor de actieve kassa." & vbCr & vbCr
    Msg = Msg & "Bent U 100 % zeker ?"
    KtrlBox = MsgBox(Msg, vbExclamation + vbYesNo + vbDefaultButton2, dokumentSleutel & " voor " & Now)
    If KtrlBox = vbNo Then Schoon: Exit Sub
    'MsgBox "Stop: Nog aan het uittesten... Enkele dagen geduld a.u.b."
        
    bBegin
    If WegBoekFout Then
        bAbort
        Exit Sub
    Else
        bEnd
        If Err Then
            MsgBox Error
            Exit Sub
        Else
            rsDetail.MoveFirst
            Do While Not rsDetail.EOF
                rsDetail("document") = dokumentSleutel
                rsDetail.MoveNext
            Loop
            SS99 Right(dokumentSleutel, 5), 11
            cmdFactuurListing_Click
            MsgBox "Kassa met succes afgesloten.  Enkel nog de ontvangst EUR " & LblIn2Btw & " (BEF " & LblInBtw & ") in kasboek inboeken a.u.b.", vbInformation
            Unload Me
        End If
    End If

End Sub

Private Sub cmdFactuurListing_Click()

dokumentSleutel = InputBox("Geef factuur of druk enter voor volledige lijst", "Kassafacturen", dokumentSleutel)
If dokumentSleutel = "" Then Exit Sub
Msg = "SELECT * FROM Kassa WHERE document Like '" & dokumentSleutel & "' ORDER BY DagUur"
Screen.MousePointer = vbHourglass
On Error GoTo 0
'Resume Next
If rsDetail.State = adStateClosed Then
Else
    rsDetail.Close
End If
Err = 0
rsDetail.Open Msg, cnnKassa, adOpenForwardOnly, adLockOptimistic
Screen.MousePointer = vbNormal
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Exit Sub
ElseIf rsDetail.RecordCount = 0 Then
    MsgBox "Er zijn (nog) geen facturen voor deze kassa die beantwoorden aan uw selectie", vbInformation, Msg
    Exit Sub
Else
    rsDetail.MoveFirst
    DrukFactuur
End If

End Sub

Private Sub cmdLijst_Click()

If cbVandaag.Value = 1 Then
    Msg = "SELECT * FROM Kassa WHERE Dag = '" + DATE_KEY(Format(Now, "dd/mm/yyyy")) + "' ORDER BY DagUur"
Else
    Msg = "SELECT * FROM Kassa ORDER BY DagUur"
End If
Screen.MousePointer = vbHourglass
On Error Resume Next
rsDetail.Close
Err = 0
rsDetail.Open Msg, cnnKassa, adOpenDynamic, adLockOptimistic
Screen.MousePointer = vbNormal
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Set rsDetail = Nothing
    Exit Sub
ElseIf rsDetail.RecordCount = 0 Then
    MsgBox "Geen verrichtingen voor deze selectie", vbInformation
    Exit Sub
End If

Unload frmADOXlog
frmADOXlog.Hide
frmADOXlog.Caption = "Kassa documenten Lijst"

With frmADOXlog.lvDetail
    ' Add ColumnHeaders. The width of the columns is
    ' the width of the control divided by the number of
    ' ColumnHeader objects.
    .ColumnHeaders.Add , , "Dag/Uur", 2000
    .ColumnHeaders.Add , , "Aantal lijnen", 1245, lvwColumnRight
    .ColumnHeaders.Add , , "EURO excl.", 1065, lvwColumnRight
    .ColumnHeaders.Add , , "EURO Btw", 1035, lvwColumnRight
    .ColumnHeaders.Add , , "BEF excl.", 1065, lvwColumnRight
    .ColumnHeaders.Add , , "BEF Btw", 1035, lvwColumnRight
    .ColumnHeaders.Add , , "Boekhouding", 3120
    ' Set View property to Report.
    .View = lvwReport
End With

Dim itmX As ListItem
Do While Not rsDetail.EOF
    ' Add a ListItem object.
    Set itmX = frmADOXlog.lvDetail.ListItems.Add(, , rsDetail("DagUur"))
        itmX.SubItems(1) = rsDetail("AantalProducten")
        itmX.SubItems(2) = rsDetail("EuroExclusief")
        itmX.SubItems(3) = rsDetail("EuroBTW")
        itmX.SubItems(4) = rsDetail("BefExclusief")
        itmX.SubItems(5) = rsDetail("BefBTW")
        itmX.SubItems(6) = rsDetail("document")
    rsDetail.MoveNext
Loop
frmADOXlog.lblAantal(0).Caption = frmADOXlog.lvDetail.ListItems.Count
Screen.MousePointer = vbNormal
frmADOXlog.cmdOptie.Caption = "Verwijderen"
frmADOXlog.cmdOptie.Tag = "Kassa"
frmADOXlog.cmdDetail.Caption = "DetailLijnen"
frmADOXlog.cmdDetail.Tag = "DetailLijnen"
GridText = ""
frmADOXlog.Show 1
frmADOXlog.cmdOptie.Tag = ""
If VSF_PRO = True And GridText <> "" Then
    Do While GridText <> ""
        Msg = Left(GridText, InStr(GridText, vbCr) - 1)
        GridText = Mid(GridText, InStr(GridText, vbCr) + 1)
        Msg = "DELETE * FROM Kassa WHERE DagUur = #" & Msg & "#"
        'SELECT * FROM Kassa IN 'c:\mijn dokumenten\marnt\data\005\kassa01.mdb' WHERE DagUur =  #06/10/00 9:47:20#
        Err = 0
        On Error Resume Next
        cnnKassa.Execute Msg, recAantal
        If Err Then
            MsgBox Err
        Else
            MsgBox "SQLopdracht: " & Msg & vbCr & vbCr & recAantal & " lijnen met succes verwijderd.", vbInformation
        End If
    Loop
ElseIf GridText <> "" Then
    MsgBox "Enkel 'gecontroleerde' verwijderingen zijn mogelijk.  Bel 053/21.59.25 voor toegangscode", vbInformation
End If

End Sub

Private Sub cmdPrintAf_Click()

    CmbAfdruk_Click
    
End Sub

Private Sub cmdSnelRekenen_Click()
Dim UitkomstHier As Variant

UitkomstHier = LineCalculating("")
If UitkomstHier Then
    MsgBox UitkomstHier, vbInformation, "Resultaat"
End If

End Sub

Private Sub CmdStock_Click()
    
CmdStock.Default = True
If KassaDetail.Enabled Then
    KassaDetail_KeyPress Asc("S")
Else
    MsgBox "Eerst kassa aanduiden a.u.b !!!", vbExclamation
End If

End Sub


Private Sub cmdSwitch_Click()
Dim TempoTel As Integer
Dim TempoVar As Variant
Dim MaskerEURBHmini As String

DIRECTSELL_STRING = cmdSwitch.Caption
MaskerEURBHmini = Mid(MASK_EURBH, 2)
If sMuntKlant = "BEF" Or sMuntKlant = "EUR" Then
'    If kassadetail.ListCount Then
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
    '    lblBTWBedrag(COUNT_TO).Caption = Format(Round(CDbl(lblBTWBedrag(COUNT_TO).Caption) * EURO, 0), "#,##0.00")
    'Next
    cmdSwitch.Caption = "Ingave in BEF"
    sMuntKlant = "BEF"
    For COUNT_TO = 0 To KassaDetail.ListCount - 1
       TempoVar = KassaDetail.List(COUNT_TO)
       If Right(TempoVar, 1) = "2" Then
       Else
           Mid(TempoVar, 42, 11) = Dec(Val(Mid(TempoVar, 42, 11)) * EURO, MaskerEURBHmini)
           Mid(TempoVar, 62, 12) = Dec(Val(Mid(TempoVar, 62, 12)) * EURO, MASK_EURBH)
           KassaDetail.List(COUNT_TO) = TempoVar
       End If
    Next
Else
    'overschakeling cijfers van BEF naar EUR
    'For COUNT_TO = 1 To 3
    '    lblBTWBedrag(COUNT_TO).Caption = Format(Round(CDbl(lblBTWBedrag(COUNT_TO).Caption) / EURO, 2), "#,##0.00")
    'Next
    cmdSwitch.Caption = "Ingave in EUR"
    sMuntKlant = "EUR"
    For COUNT_TO = 0 To KassaDetail.ListCount - 1
        TempoVar = KassaDetail.List(COUNT_TO)
        If Right(TempoVar, 1) = "2" Then
        Else
            Mid(TempoVar, 42, 11) = Dec(Val(Mid(TempoVar, 42, 11)) / EURO, MaskerEURBHmini)
            Mid(TempoVar, 62, 12) = Dec(Val(Mid(TempoVar, 62, 12)) / EURO, MASK_EURBH)
            KassaDetail.List(COUNT_TO) = TempoVar
        End If
    Next
End If
'TekstInfo(5).Text = Dec$(dMuntK, "##0.########")
DIRECTSELL_STRING = cmdSwitch.Caption
RefreshBTW

End Sub

Private Sub cmdWissel_Click()

    'Unload Wissel
    'Wissel.Show 1
    
End Sub

Private Sub cmdWissel2_Click()

    'Unload Wissel
    'Wissel.Show 1
    'If KasBetalingBEF + KasBetalingEUR + KasTerugEUR = 0 Then
    'Else
    '    If KasBetalingBEF Then
    '        obBetaling(5).Value = True
    '        txtOntvangenBEF.Text = Str$(KasBetalingBEF)
    '        txtOntvangenEUR.Text = "0"
    '    Else
    '        obBetaling(0).Value = True
    '        txtOntvangenBEF.Text = "0"
    '        txtOntvangenEUR.Text = Str$(KasBetalingEUR)
    '    End If
    '    txtTeruggaveEUR.Text = Str$(KasTerugEUR)
    'End If
    
End Sub


Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If
VAT_BOBTHEBUILDERS = False

Dim Tel As Integer

Screen.MousePointer = vbHourglass
For Tel = 16 To 19
    rbtwVAK(Tel - 16) = String99(READING, Tel)
    rbtwVAK(Tel - 12) = String99(READING, Tel + 6)
Next
Schoon
cbVandaag.Caption = "&Enkel " & Format(Now, "dd/mm/yyyy")
cmdFactuurListing.Visible = True
On Error Resume Next
cbEpsonTM.Value = LaadTekst("KassaVerkoop", "EpsonTMModus")
Screen.MousePointer = vbNormal
If bhEuro Then
    DIRECTSELL_STRING = "Ingave in EUR"
Else
    DIRECTSELL_STRING = "Ingave in BEF"
End If

If bhEuro Then
    cmdSwitch.Caption = "Ingave in EUR"
    sMuntKlant = "EUR"
Else
    cmdSwitch.Caption = "Ingave in BEF"
    sMuntKlant = "BEF"
End If

    KasTicketTotaal = 0
    KasTotaal = 0
    KasBEF = 0
    KasEUR = 0
    KasTotaalBEF = 0
    KasTotaalEUR = 0

End Sub


Private Sub InstalKlant()
Dim Klantje As String
Dim T       As Integer

On Local Error Resume Next

KassaDetail.Enabled = True
cmdFactuurListing.Enabled = True

'InvoerLijn.Enabled = True
Klantje = vbCrLf + vBibTekst(TABLE_CUSTOMERS, "#A100 #") + vbCrLf + vBibTekst(TABLE_CUSTOMERS, "#A125 #") + vbCrLf + vBibTekst(TABLE_CUSTOMERS, "#A104 #") + vbCrLf + vBibTekst(TABLE_CUSTOMERS, "#A109 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A107 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A108 #")
KassaNaam = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
Schoonvegen.Enabled = True
'ReturnRooster.Enabled = True

sMunt = Mim.SnelHelp.Panels(2)
'If sMunt = "BEF" Then
    dMuntK = 1
'Else
'    dMuntK = EURO
'End If

If vBibTekst(TABLE_CUSTOMERS, "#v149 #") = "" Then
    MsgBox "Landnummer is verplicht !"
    Exit Sub
ElseIf vBibTekst(TABLE_CUSTOMERS, "#v149 #") = "002" Then
    VerkoopFLG = 0
    frmKassaVerkoop.Caption = frmKassaVerkoop.Caption + " - " + vBibTekst(TABLE_CUSTOMERS, "#A100 #") + " (" + vBibTekst(TABLE_CUSTOMERS, "#A110 #") + ")"
    'kontroleren LU speciaal
Else
    MsgBox "Enkel binnenland..."
    Exit Sub
End If

If Left(vBibTekst(TABLE_CUSTOMERS, "#v161 #"), 3) = "400" Then
    KlantRekening = vBibTekst(TABLE_CUSTOMERS, "#v161 #")
    bGet TABLE_LEDGERACCOUNTS, 0, KlantRekening
    If Ktrl Then
        Beep
        MsgBox KlantRekening & "-Klantenrekening bestaat niet", vbCritical
        KlantRekening = String99(READING, 9)
    End If
Else
    MsgBox "400xxxx-KlantenSUBrekening (aanbevolen) aanduiden voor " + vBibTekst(TABLE_CUSTOMERS, "#A110 #"), vbExclamation
    KlantRekening = String99(READING, 9)
End If
If Left(vBibTekst(TABLE_CUSTOMERS, "#v225 #"), 2) = "70" Then
    DefaultVerkoop = vBibTekst(TABLE_CUSTOMERS, "#v225 #")
    bGet TABLE_LEDGERACCOUNTS, 0, DefaultVerkoop
    If Ktrl Then
        Beep
        DefaultVerkoop = String99(READING, 25)
    End If
Else
    DefaultVerkoop = String99(READING, 25)
End If
If Dir(LOCATION_COMPANYDATA + "\" + Trim$(vBibTekst(TABLE_CUSTOMERS, "#A110 #")) + ".mdb") = "" Then
    MsgBox LOCATION_COMPANYDATA + Trim$(vBibTekst(TABLE_CUSTOMERS, "#A110 #")) + ".mdb wordt eerst aangemaakt", vbInformation
    If Not CopyFile(App.path, LOCATION_COMPANYDATA, "Kassa.mdb") Then
        MsgBox "Fout tijdens aanmaak " + LOCATION_COMPANYDATA + Trim$(vBibTekst(TABLE_CUSTOMERS, "#A110 #")) + ".mdb" + vbCr + vbCr + "Venster kassaverkoop wordt hierna gesloten", vbCritical
        Exit Sub
    Else
        Dim OldName As String
        Dim NewName As String
        
        OldName = LOCATION_COMPANYDATA + "Kassa.mdb"
        NewName = LOCATION_COMPANYDATA + Trim$(vBibTekst(TABLE_CUSTOMERS, "#A110 #")) + ".mdb"
        Name OldName As NewName   ' Move and rename file.
        MsgBox LOCATION_COMPANYDATA + Trim$(vBibTekst(TABLE_CUSTOMERS, "#A110 #")) + ".mdb met succes aangemaakt.", vbInformation
    End If
End If
  
Set cnnKassa = New ADODB.Connection
Set rsDetail = New ADODB.Recordset

Dim KassaConnect As String

KassaConnect = ADOJET_PROVIDER & _
    "Data Source=" & LOCATION_COMPANYDATA & _
    Trim$(vBibTekst(TABLE_CUSTOMERS, "#A110 #")) + ".mdb;" & _
    "Persist Security Info=False"

On Error Resume Next
Err = 0
cnnKassa.Open KassaConnect
If Err Then MsgBox Error: Exit Sub
rsDetail.CursorLocation = adUseClient
X = TellerKassa

End Sub

Private Sub cmdKassaKiezen_Click()

If KassaDetail.ListCount Then
          Msg = "Huidige inbreng negeren." + vbCrLf + vbCrLf
    Msg = Msg + "Bent U zeker ?"
    KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
    If KtrlBox = vbYes Then
    Else
        Exit Sub
    End If
End If

aIndex = 1
SharedFl = TABLE_CUSTOMERS
GridText = "KASSA"
SqlSearch.Show 1
If Ktrl = 0 Then
    If UCase(Left(XLogKey, 5)) <> "KASSA" Then
        MsgBox "Kassa's dienen VERPLICHT klantenkode startend met 'Kassa' te hebben", vbExclamation
    Else
        InstalKlant
    End If
Else
    Schoon
End If

End Sub

Private Sub MaakTotaal()
Dim BTWIn As Double
Dim BTWEx As Double

Dim BtwEuroIn As Double
Dim BtwEuroEx As Double

Dim Teller As Integer

On Local Error Resume Next

For Teller = 1 To 3
    BTWBedrag(Teller) = Val(Dec(BTWBasis(Teller) * Val(Mid(fmarBoxText("002", "2", Format(Teller)), 4, 4)) / 100, MASK_EURBH))
    BTWIn = BTWIn + BTWBasis(Teller) + BTWBedrag(Teller)
    BTWEx = BTWEx + BTWBasis(Teller)

    BTWEuroBedrag(Teller) = Val(Dec(BTWEuroBasis(Teller) * Val(Mid(fmarBoxText("002", "2", Format(Teller)), 4, 4)) / 100, MASK_EURBH))
    BtwEuroIn = BtwEuroIn + BTWEuroBasis(Teller) + BTWEuroBedrag(Teller)
    BtwEuroEx = BtwEuroEx + BTWEuroBasis(Teller)
    If cmdSwitch.Caption = "Ingave in BEF" Then
        'lblBTWBedrag(Teller) = Format(Round(BTWBasis(Teller)), "#,##0.00")
    Else
        'lblBTWBedrag(Teller) = Format(BTWEuroBasis(Teller), "#,##0.00")
    End If
Next
BTWEx = BTWEx + BTWBasis(0)
BTWIn = BTWIn + BTWBasis(0)

BtwEuroEx = BtwEuroEx + BTWEuroBasis(0)
BtwEuroIn = BtwEuroIn + BTWEuroBasis(0)

LblExBtw = Format(Round(BTWEx, 2), "#,##0.00")   'in BEF
LblInBtw = Format(Round(BTWIn), "#,##0.00")    'in BEF
LblEx2Btw = Format(BtwEuroEx, "#,##0.00") 'masker voor EURO
LblIn2Btw = Format(BtwEuroIn, "#,##0.00")  'masker voor EURO
LblEurCash = Format(Round(BtwEuroIn / 5, 2) * 5, "#,##0.00")

End Sub

Private Sub RasterSchoon()

KassaDetail.Clear
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

For T = 0 To KassaDetail.ListCount - 1
    aa = KassaDetail.List(T)
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
                        BTWBasis(bVak) = Val(Dec(BTWBasis(bVak), maskerMULTI))
                    End If
                    BTWBasis(bVak) = BTWBasis(Val(Mid(aa, 88, 1))) + Val(Dec(Val(Mid(aa, 62, 12)), maskerMULTI))
                    BTWEuroBasis(bVak) = BTWEuroBasis(Val(Mid(aa, 88, 1))) + Val(Dec(Val(Mid(aa, 62, 12)) / EURO, MASK_EURBH))
                End If
            End If
    End Select
Next
MaakTotaal

End Sub

Private Sub Schoon()
Dim T As Integer

'KassaDetail.Enabled = False
Afsluiten.Enabled = False

CmbAfdruk.Enabled = False '
LblExBtw = ""
LblInBtw = ""
LblEx2Btw = ""
LblIn2Btw = ""
LblEurCash = ""

txtOntvangenBEF.text = "0"
txtOntvangenEUR.text = "0.00"
txtTeruggaveEUR.text = "0.00"

RasterSchoon
frmKassaVerkoop.SSTab1.Tab = 0
obBetaling(0).Value = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

If cmdKassaKiezen.Enabled = False Then
    On Error Resume Next
    Err = 0
    rsDetail.Close
    If Err Then MsgBox Error: Exit Sub
    cnnKassa.Close
    If Err Then MsgBox Error: Exit Sub
    Set rsDetail = Nothing
    Set cnnKassa = Nothing
End If

End Sub

Private Sub LblIn2Cash_Click(Index As Integer)

End Sub

Private Sub obBetaling_Click(Index As Integer)

    If Index = 0 Or Index = 5 Then
        fContant.Visible = True
    Else
        fContant.Visible = False
    End If
    If Index = 1 Or Index = 2 Then
        'MsgBox "Banksys Drivers installeren a.u.b.", vbInformation
        'obBetaling(0).Value = True
    End If
    
End Sub

Private Sub SchoonVegen_Click()

Schoon

End Sub



Private Sub KassaDetail_DblClick()

If KassaDetail.ListIndex = -1 Then
    KassaDetail_KeyPress 43
Else
    KassaDetail_KeyPress 13
End If

End Sub

Private Sub KassaDetail_GotFocus()

If CmdStock.Default Then CmdStock.Default = False
SnelHelpPrint "Alt+[S]tock: bijvoegen, Insert: plakken, Delete: knippen of verwijderen, Enter: wijzigen", BL_LOGGING

End Sub


Private Sub KassaDetail_KeyPress(KeyAscii As Integer)
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
Positie = KassaDetail.ListIndex
Select Case KeyAscii
    Case 13         'Enter
        If Positie < 0 Then
            MsgBox "Eerst een lijn selekteren !", 0, "Lijn wijzigen"
            Exit Sub
        End If
        If dokumentType <> "15" And Annuleren.Enabled = False Then
            Annuleren.Enabled = True
        End If
        GridText = KassaDetail.text
        If Right(GridText, 1) = "2" Then
        Else
            RefreshBTW
        End If
        XLogKassa = ""
        WijzigenVerkoop.Show 1
        Unload WijzigenVerkoop
        frmKassaVerkoop.SetFocus
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
            KassaDetail.RemoveItem Positie
            KassaDetail.AddItem GridText, Positie
        End If
        KassaDetail.ListIndex = Positie
        MaakTotaal

    Case 43
        GridText = ""
        XLogKassa = ""
        WijzigenVerkoop.Show 1
        Unload WijzigenVerkoop
        frmKassaVerkoop.SetFocus
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
                KassaDetail.AddItem GridText, KassaDetail.ListCount
                If blMilieu = True Then
                    splitmilieu = Split(MilieuGridText, vbCrLf)
                    For telmilieu = 0 To UBound(splitmilieu) - 1
                        KassaDetail.AddItem splitmilieu(telmilieu), KassaDetail.ListCount
                    Next
                    blMilieu = False
                End If
            Else
                KassaDetail.AddItem GridText, Positie
                If blMilieu = True Then
                    MsgBox "Stop, waarom?"
                    splitmilieu = Split(MilieuGridText, vbCrLf)
                    For telmilieu = 0 To UBound(splitmilieu) - 1
                        KassaDetail.AddItem splitmilieu(telmilieu), KassaDetail.ListCount
                    Next
                    blMilieu = False
                End If
            End If
        End If
        If Mid(XLogKey, 39, 2) = vbCrLf Then
            Do While Len(XLogKey) <> 0
                Msg = Mid(XLogKey, 1, 38) + String(37, " ") + "|2"
                KassaDetail.AddItem Msg
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
If KassaDetail.ListCount Then
    Afsluiten.Enabled = True
    CmbAfdruk.Enabled = True
Else
    Afsluiten.Enabled = False
    CmbAfdruk.Enabled = False
End If
RefreshBTW
If CmdStock.Default Then CmdStock.SetFocus
End Sub

Private Sub KassaDetail_LostFocus()

KassaDetail.ListIndex = -1

End Sub



Private Function TellerKassa()

Dim TestString As String

JUMPERtellerKASSA:
Msg = "SELECT * FROM Kassa WHERE Dag = '" + DATE_KEY(Format(Now, "dd/mm/yyyy")) + "'"
Screen.MousePointer = vbHourglass
On Error Resume Next
If rsDetail.State = adStateClosed Then
Else
    rsDetail.Close
End If
Err = 0
rsDetail.Open Msg, cnnKassa, adOpenForwardOnly, adLockOptimistic
Screen.MousePointer = vbNormal
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Set rsDetail = Nothing
    Exit Function
ElseIf rsDetail.RecordCount = 0 Then
    'test
    On Error Resume Next
    Err = 0
    'ALTER TABLE eShop ADD COLUMN BetaaldEUR MONEY
    TestString = rsDetail("BetaaldEUR").Name
    If Err Then
        MsgBox "Bestaande Kassadatabase wordt eerst bijgewerkt met nieuwe velden", vbInformation
        rsDetail.Close
        Err = 0
        cnnKassa.Execute "ALTER TABLE Kassa ADD COLUMN BetaaldEUR MONEY"
        cnnKassa.Execute "ALTER TABLE Kassa ADD COLUMN BetaaldBEF MONEY"
        cnnKassa.Execute "ALTER TABLE Kassa ADD COLUMN TerugEUR MONEY"
        cnnKassa.Execute "ALTER TABLE Kassa ADD COLUMN BetaalType TEXT(25)"
        GoTo JUMPERtellerKASSA
    Else
        MsgBox "Klaar voor eerste verrichting vandaag", vbInformation
    End If
Else
    rsDetail.MoveLast
    If IsNull(rsDetail("document")) Then
        cmdDagAfsluiting.Enabled = True
        cmdDagAfsluiting.Caption = "&Dag Afsluiting voor " & Format(Now, "dd/mm/yyyy")
    Else
        MsgBox Format(Now, "dd/mm/yyyy") & " is reeds geboekt (U kan enkel controleren via totaallijst)", vbExclamation
        cmdKassaKiezen.Enabled = False
        CmdStock.Enabled = False
        cmdLijst.Enabled = True
        Schoonvegen.Enabled = False
        Exit Function
    End If
End If
StatusBar1.Panels(3).text = Format(rsDetail.RecordCount + 1, "0000")
cmdKassaKiezen.Enabled = False
cmdLijst.Enabled = True
frmKassaVerkoop.SetFocus

End Function


Private Function WegBoekFout%()
Dim T               As Integer
Dim AantalStuks     As Single
Dim KassaNummer     As String
Dim KassaNaam       As String

Dim dKtrlBTW        As Currency
Dim KtrlVierkant As Currency

WegBoekFout = True
DKTRL_CUMUL = 0: DKTRL_BEF = 0: DKTRL_EUR = 0
Unload frmBoeking
Load frmBoeking

TLB_RECORD(TABLE_INVOICES) = ""
'bijgevoegd voor scanning..
If rsMAR(TABLE_INVOICES).State = adStateClosed Then
    Ktrl = bOpen(TABLE_INVOICES)
End If
rsMAR(TABLE_INVOICES).AddNew

KassaNummer = Mid(frmKassaVerkoop.Caption, InStr(frmKassaVerkoop.Caption, "(") + 1)
KassaNummer = Mid(KassaNummer, 1, Len(KassaNummer) - 1)
bGet TABLE_CUSTOMERS, 0, KassaNummer
If Ktrl Then
    MsgBox "Kassa niet toegankelijk..."
    Exit Function
Else
    RecordToVeld TABLE_CUSTOMERS
End If
vBib TABLE_INVOICES, dokumentSleutel, "v033"
vBib TABLE_INVOICES, "K" + rsMAR(TABLE_CUSTOMERS)("A110"), "v034"
vBib TABLE_INVOICES, DATE_KEY(Format(Now, "dd/mm/yyyy")), "v035"
vBib TABLE_INVOICES, DATE_KEY(Format(Now, "dd/mm/yyyy")), "v036"
vBib TABLE_INVOICES, Dec$(dMuntK, "###.0000000000"), "v040"
vBib TABLE_INVOICES, rsMAR(TABLE_CUSTOMERS)("vs03"), "vs03"

If BTWBasis(0) <> 0 Then
    If bhEuro Then
        vBib TABLE_INVOICES, Str$(BTWEuroBasis(0)), "v055"
    Else
        vBib TABLE_INVOICES, Format(BTWBasis(0), MASK_SY(0)), "v055"
    End If
End If

TotaalBTW = 0
For T = 1 To 3
    If bhEuro Then
        If BTWEuroBedrag(T) <> 0 Then
            vBib TABLE_INVOICES, Str$(BTWEuroBasis(T)), "v" + Format(55 + T, "000")
            TotaalBTW = TotaalBTW + BTWEuroBedrag(T)
        End If
    Else
        If BTWBedrag(T) <> 0 Then
            vBib TABLE_INVOICES, Format(BTWBasis(T), MASK_SY(0)), "v" + Format(55 + T, "000")
            TotaalBTW = TotaalBTW + Val(Dec(BTWBedrag(T), MASK_SY(0)))
        End If
    End If
Next

If bhEuro Then
    KtrlVierkant = CDbl(LblIn2Btw) - CDbl(LblEx2Btw)
Else
    KtrlVierkant = CDbl(LblInBtw) - CDbl(LblExBtw)
End If

If KtrlVierkant <> TotaalBTW Then
    MsgBox "Vierkantskontrole BTW: " & KtrlVierkant & " <> " & TotaalBTW & vbCr & vbCr & "De boeking wordt niet uitgevoerd.  Kontakteer ons 053/21.59.25", vbExclamation
    Schoon
    Exit Function
End If
If bhEuro Then
    vBib TABLE_INVOICES, Str$(TotaalBTW), "v064"
Else
    vBib TABLE_INVOICES, Format(TotaalBTW, MASK_SY(0)), "v064"
End If

If ForFait = 1 Then
    GoTo JumpToKassaDoc
End If

For T = 0 To KassaDetail.ListCount - 1
    If -Val(Mid(KassaDetail.List(T), 62, 12)) = 0 Then
    Else
        GoSub JournaalNew
        rsJournaal("v067") = Trim(Mid(KassaDetail.List(T), 1, 30))
        rsJournaal("v102") = Trim(Mid(KassaDetail.List(T), 98, 13))
        rsJournaal("v068") = Dec(-Val(Mid(KassaDetail.List(T), 62, 12)), "#.00")
        'vroeger rsJournaal("v068") = Str$(-Val(Mid(KassaDetail.List(T), 62, 12)))
        rsJournaal("dece068") = Val(Dec(-Val(Mid(KassaDetail.List(T), 62, 12)), "#.00"))
        'vroeger rsJournaal("dece068") = -Val(Mid(KassaDetail.List(T), 62, 12))
        rsJournaal("v019") = Mid(KassaDetail.List(T), 90, 7)
        If Not ADOJOURNAL_OK Then
            MsgBox Error
            DKTRL_CUMUL = 99
            GoTo VKKontrole
        End If
    End If
Next
      
If TotaalBTW Then
    GoSub JournaalNew
    rsJournaal("v067") = rsMAR(TABLE_CUSTOMERS)("A100")
    rsJournaal("v068") = Str$(-TotaalBTW)
    rsJournaal("dece068") = -TotaalBTW
    rsJournaal("v019") = rbtwVAK(0)
    If Not ADOJOURNAL_OK Then
        MsgBox Error
        DKTRL_CUMUL = 99
        GoTo VKKontrole
    End If
End If

GoSub JournaalNew
rsJournaal("v067") = rsMAR(TABLE_CUSTOMERS)("A100")
rsJournaal("v019") = KlantRekening
If bhEuro Then
    KtrlVierkant = CDbl(LblEx2Btw) + TotaalBTW
Else
    KtrlVierkant = CDbl(LblExBtw) + TotaalBTW
End If
If -DKTRL_CUMUL <> KtrlVierkant Then
    If Abs(-DKTRL_CUMUL - KtrlVierkant) < 0.05 Then
        Msg = "Rekenverschil " & -DKTRL_CUMUL & " <> " & KtrlVierkant & vbCr & vbCr
        Msg = Msg & "Totaal dagfactuur wordt geboekt als " & -DKTRL_CUMUL
        MsgBox Msg, vbExclamation
        rsJournaal("v068") = Str$(-DKTRL_CUMUL)
        rsJournaal("dece068") = -DKTRL_CUMUL
    Else
        MsgBox "Rekenverschil groter dan EUR 0.05 voor één dag", vbExclamation
        DKTRL_CUMUL = 99
        GoTo VKKontrole
    End If
Else
    rsJournaal("v068") = Str$(KtrlVierkant)
    rsJournaal("dece068") = KtrlVierkant
End If
If Not ADOJOURNAL_OK Then
    MsgBox Error
    DKTRL_CUMUL = 99
    GoTo VKKontrole
End If

JumpToKassaDoc:
vBib TABLE_INVOICES, Str$(KtrlVierkant), "v249"
bInsert TABLE_INVOICES, 0

VKKontrole:
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

    If DKTRL_CUMUL Then
        WegBoekFout = True
    Else
        WegBoekFout = False
    End If
End If
Exit Function

FoutMelding:
MsgBox "Fout tijdens het wegboeken." + vbCrLf + "Hierna worden alle bewerkingen genegeerd." + vbCrLf + vbCrLf + "Begin opnieuw."
WegBoekFout = True
Exit Function

JournaalNew:
rsJournaal.AddNew
rsJournaal("v034") = "K" + rsMAR(TABLE_CUSTOMERS)("A110")
rsJournaal("v066") = DATE_KEY(Format(Now, "dd/mm/yyyy"))
rsJournaal("v035") = DATE_KEY(Format(Now, "dd/mm/yyyy"))
rsJournaal("v033") = dokumentSleutel
rsJournaal("v069") = KlantRekening
Return

End Function


Function DrukFactuur()

On Error GoTo 0

Dim varVar As Variant
Dim Veldjes As Variant
Dim dokumentNummer As String
Dim rsdokument As New ADODB.Recordset
Dim rsKlant As New ADODB.Recordset
Dim PAGE_COUNTER As Long
Dim FactuurDatum As String
Dim NaamKassa As String * 40
Dim TabNu As Boolean

Dim TotaalEUR As Currency
Dim TotaalBEF As Currency
Dim Vak00 As Currency
Dim Vak01 As Currency
Dim Vak02 As Currency
Dim Vak03 As Currency
Dim TotaalBTW As Currency

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If Printer.Orientation = vbPRORLandscape Then
    Printer.Orientation = vbPRORPortrait
End If
DoEvents
Printer.FontName = "courier new"

dokumentNummer = ""
Do While Not rsDetail.EOF
    If dokumentNummer = "" Then
        dokumentNummer = rsDetail("document")
        PAGE_COUNTER = 0
        GoSub PrintKopTekst
    ElseIf dokumentNummer <> rsDetail("document") Then
        GoSub PrintVoetTekst
        dokumentNummer = rsDetail("document")
        PAGE_COUNTER = 0
        GoSub PrintKopTekst
    End If
    varVar = Split(rsDetail(7).Value, vbCr)
    For COUNT_TO = 0 To UBound(varVar) - 1
        Veldjes = Split(varVar(COUNT_TO), Chr(124))
        GoSub PrintMiddenTekst
    Next
    rsDetail.MoveNext
Loop
GoSub PrintVoetTekst
Exit Function

PrintKopTekst:
PAGE_COUNTER = PAGE_COUNTER + 1
rsdokument.CursorLocation = adUseClient
rsdokument.Open "SELECT * FROM dokumenten WHERE v033 ='" & dokumentNummer & "'", adntDB, adOpenForwardOnly, adLockReadOnly
If rsdokument.RecordCount Then
    'dokumentNummer
    FactuurDatum = DATE_TEXT(rsdokument("v035"))
    rsKlant.CursorLocation = adUseClient
    rsKlant.Open "SELECT * FROM Klanten WHERE A110 ='" & Mid(rsdokument("v034"), 2) & "'", adntDB, adOpenForwardOnly, adLockReadOnly
    If rsKlant.RecordCount Then
        NaamKassa = Trim$(Mid(rsdokument("v034"), 2)) & " " & rsKlant("A100")
    Else
        NaamKassa = Trim$(Mid(rsdokument("v034"), 2)) & " " & "Kassafiche niet gevonden..."
    End If
    rsKlant.Close
    'Koptekst Lijn 1
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Print FULL_LINE
    Printer.Print "KassaID " & String(33, " ") & "Blz. " & "document    " _
                & "Datum      " & "Betalingstuk"
    Printer.FontBold = True
    Printer.Print NaamKassa & "  " & Dec((PAGE_COUNTER), "###") & " " & rsdokument("v033") _
                & " " & FactuurDatum & " " & rsdokument("v038")
    Printer.FontBold = False
    Printer.Print FULL_LINE
       
    Vak00 = Val(rsdokument("v055"))
    Vak01 = Val(rsdokument("v056"))
    Vak02 = Val(rsdokument("v057"))
    Vak03 = Val(rsdokument("v058"))
    TotaalBTW = Val(rsdokument("v064"))
    TotaalBEF = Vak00 + Vak01 + Vak02 + Vak03 + TotaalBTW
    'Koptekst Lijn 2
    Printer.Print "    VAK 00 " & "    VAK 01 " & "    VAK 02 " & "    VAK 03 " & "       Btw " & "    Totaal"
    Printer.FontBold = True
    Printer.Print Dec((Vak00), MASK_EUR) & " " & Dec((Vak01), MASK_EUR) & " " & Dec((Vak02), MASK_EUR) & " " & Dec((Vak03), MASK_EUR) & " " & Dec((TotaalBTW), MASK_EUR) & " " & Dec((TotaalBEF), MASK_EUR)
    Printer.FontBold = False
    Printer.Print FULL_LINE
    Printer.FontSize = 8
    Printer.Print "Artikel Code " & " " & "Omschrijving                   " & "    excl.  " & "Artikel Code " & " " & "Omschrijving                   " & "BEF excl."
    Printer.Print FULL_LINE
    TabNu = False
Else
    MsgBox "Stop"
End If
rsdokument.Close
Return

PrintMiddenTekst:
TabNu = Not TabNu
If TabNu Then Printer.Print vbCrLf;
Printer.Print Veldjes(9) & " " & Mid(Veldjes(0), 1, 30) & " " & Dec(Val(Veldjes(3)), MASK_EUR); " ";
Return

PrintVoetTekst:
Printer.EndDoc
Return
    
End Function


Private Sub SSTab1_Click(PreviousTab As Integer)

    If SSTab1.Caption = "&Ticket" Then
        Screen.MousePointer = vbHourglass
        txtTicket.text = ""
        If KassaDetail.ListCount > 0 Then
            Ktrl = PrintAf(False, True)
            txtTicket.SetFocus
            cmdPrintAf.Enabled = True
        Else
            cmdPrintAf.Enabled = False
        End If
        Screen.MousePointer = vbNormal
    ElseIf SSTab1.Caption = "&Betalen" Then
        If KassaDetail.ListCount = 0 Then
            cbVerwerken2.Enabled = False
        Else
            cbVerwerken2.Enabled = True
        End If
        'KasTicketTotaal = CDbl(LblIn2Btw.Caption)
        KasTicketTotaal = CDbl(LblEurCash.Caption)
        txtTicketTotaal.text = Format(KasTicketTotaal, "€ #,##0.00")
        obBetaling(0).Value = True
        txtOntvangenEUR.text = Str$(KasTicketTotaal)
        txtOntvangenBEF.text = "0"
    End If

End Sub

Private Sub txtOntvangenBEF_Change()

    Dim HierBEF As Currency
    Dim HierEUR As Currency
    Dim HierTotaal As Currency
    
    obBetaling(5).Value = True
    HierEUR = Val(txtOntvangenEUR.text)
    HierBEF = Round(Val(txtOntvangenBEF.text) / EURO, 2)
    If txtTicketTotaal.text = "" Then
        txtTicketTotaal.text = "€ 0"
    End If
    HierTotaal = CDbl(Mid(txtTicketTotaal.text, 2)) - HierEUR - HierBEF
    txtTeruggaveEUR.text = Str$(-HierTotaal)
    'If HierTotaal < 0 Then
    '    txtTeruggaveEUR.BackColor = &HFF&
    'Else
    '    txtTeruggaveEUR.BackColor = &H80000005
    'End If
    
End Sub

Private Sub txtOntvangenEUR_Change()
    
    Dim HierBEF As Currency
    Dim HierEUR As Currency
    Dim HierTotaal As Currency
    
    obBetaling(0).Value = True
    HierEUR = Val(txtOntvangenEUR.text)
    'HierBEF = Round(Val(txtOntvangenBEF.Text) / EURO, 2)
    If txtTicketTotaal.text = "" Then
        txtTicketTotaal.text = "€ 0"
    End If
    HierTotaal = CDbl(Mid(txtTicketTotaal.text, 2)) - HierEUR '- HierBEF
    txtTeruggaveEUR.text = Str$(-HierTotaal)
    'If HierTotaal < 0 Then
    '    Stop
    'Else
    '    Stop
    'End If
    
End Sub

Private Sub txtTeruggaveEUR_Change()

    txtTeruggaveBEF.text = Format(Round(Val(txtTeruggaveEUR.text) * EURO), "#,##0")
    
End Sub

Function BoekingKassa() As Boolean

    BoekingKassa = False
    On Error Resume Next
    Err = 0
    bBegin
    cnnKassa.BeginTrans
    If WegSchrijfFout Then
        cnnKassa.RollbackTrans
        bAbort
        Exit Function
    Else
        cnnKassa.CommitTrans
        bEnd
        If Err Then
            MsgBox Error
            Exit Function
        Else
            X = TellerKassa
            Schoon
        End If
    End If
    If Err = 0 Then
        BoekingKassa = True
    End If

End Function

