VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0E400164-0685-434B-B5FF-CE76FE37CF14}#6.10#0"; "vpectrl61.ocx"
Begin VB.MDIForm Mim 
   BackColor       =   &H8000000C&
   ClientHeight    =   9600
   ClientLeft      =   4935
   ClientTop       =   2175
   ClientWidth     =   14070
   Icon            =   "frmMIM.frx":0000
   NegotiateToolbars=   0   'False
   Visible         =   0   'False
   Begin MSComDlg.CommonDialog Teken 
      Left            =   5400
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   5760
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin MSComctlLib.StatusBar SnelHelp 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   9315
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   2
            Object.ToolTipText     =   "Bedrijfstype"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   883
            MinWidth        =   883
            Text            =   "--"
            TextSave        =   "--"
            Object.ToolTipText     =   "Munt van het boekjaar"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   883
            MinWidth        =   883
            Text            =   "JET4"
            TextSave        =   "JET4"
            Object.ToolTipText     =   "OLEDB driver, SQL Server of Jet (Access compatibel)"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
            Object.ToolTipText     =   "Hier komen begeleidende info en/of gebruiktips"
         EndProperty
      EndProperty
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
   Begin VB.Timer SnelHelpTijd 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   8760
      Top             =   2640
   End
   Begin MSComctlLib.Toolbar tbToolBar 
      Align           =   1  'Align Top
      Height          =   750
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   14070
      _ExtentX        =   24818
      _ExtentY        =   1323
      ButtonWidth     =   1852
      ButtonHeight    =   1164
      Appearance      =   1
      ImageList       =   "imlIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open"
            Key             =   "Open"
            Object.ToolTipText     =   "Bedrijf Openen"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SQL"
            Key             =   "SQLZoek"
            Object.ToolTipText     =   "Aktieve Bedrijfsdatabase bewerken via SQL"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Map Manueel"
            Key             =   "vsoft"
            Object.ToolTipText     =   "Klantdocumenten In/Uit map (serverlocatie)"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Archief"
            Key             =   "server2"
            Object.ToolTipText     =   "Archiefmap (serverlocatie)"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.CommandButton cbMovie3 
         Caption         =   "InterVat?"
         Enabled         =   0   'False
         Height          =   375
         Left            =   11040
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton cbMovie2 
         Caption         =   "KoppelKost?"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8520
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cbMovie1 
         Caption         =   "Demo dwLd"
         Enabled         =   0   'False
         Height          =   375
         Left            =   9720
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.ComboBox cmbKassaPrinter 
         Enabled         =   0   'False
         Height          =   315
         Left            =   12120
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Kassa Printer"
         Top             =   0
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DatumVerwerking 
         Height          =   315
         Left            =   12480
         TabIndex        =   7
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   77070339
         CurrentDate     =   39083
         MaxDate         =   58862
         MinDate         =   31168
      End
      Begin VB.ComboBox CmbLijstPrinter 
         Enabled         =   0   'False
         Height          =   315
         Left            =   8520
         Style           =   2  'Dropdown List
         TabIndex        =   1
         ToolTipText     =   "Lijsten Printer"
         Top             =   0
         Width           =   3495
      End
      Begin VB.ComboBox CmbDokumentPrinter 
         Enabled         =   0   'False
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   0
         ToolTipText     =   "Dokumenten printer"
         Top             =   0
         Width           =   3675
      End
      Begin MSMask.MaskEdBox FontGrootte 
         Height          =   225
         Left            =   7440
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   397
         _Version        =   393216
         AutoTab         =   -1  'True
         PromptChar      =   "_"
      End
      Begin VB.ComboBox cmdWegBoekModus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMIM.frx":030A
         Left            =   4800
         List            =   "frmMIM.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Maak uw keuze en druk ENTER om te vergrendelen (ontgrendelen is mogelijk via Akties)"
         Top             =   360
         Width           =   3675
      End
      Begin VB.ComboBox cmbpaperbin 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   3915
      End
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   7440
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMIM.frx":030E
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMIM.frx":042E
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMIM.frx":0710
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMIM.frx":0B64
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMIM.frx":0FB6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMIM.frx":1408
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox InfoData 
      Align           =   3  'Align Left
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000018&
      Height          =   8565
      Left            =   0
      ScaleHeight     =   8505
      ScaleWidth      =   3660
      TabIndex        =   3
      Top             =   750
      Visible         =   0   'False
      Width           =   3720
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
         Height          =   1740
         Left            =   -120
         TabIndex        =   4
         Top             =   3000
         Visible         =   0   'False
         Width           =   8055
      End
      Begin VPECTRLLib.VPE Report 
         Height          =   1155
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   3495
         _Version        =   393226
         Appearance      =   0
         Caption         =   "VPE Preview"
         Enabled         =   -1  'True
         PageWidth       =   21
         PageHeight      =   29.7
         SwapFileName    =   ""
         RulersMeasure   =   0
         ExternalWindow  =   -1  'True
         GridMode        =   0
         tbGrid          =   0   'False
         Rulers          =   -1  'True
         ToolBar         =   -1  'True
         tbClose         =   -1  'True
         tbScale         =   -1  'True
         StatusBar       =   -1  'True
         tbOpen          =   -1  'True
         tbSave          =   -1  'True
         tbPrint         =   -1  'True
         tbHelp          =   -1  'True
         tbAbout         =   -1  'True
         tbNavigate      =   -1  'True
         GridVisible     =   0   'False
         PageFormat      =   -1
         PaperView       =   -1  'True
         PageScroller    =   -1  'True
         StatusSegment   =   -1  'True
         EnableURLs      =   -1  'True
         PageScrollerTracking=   -1  'True
         PreviewCtrl     =   1
         DocFileReadOnly =   0   'False
         EnableHelpRouting=   0   'False
         EnablePrintSetupDialog=   -1  'True
         tbMail          =   -1  'True
         PageOrientation =   1
         PictureCacheSize=   65536
      End
      Begin VB.Image imgFiguur 
         Height          =   1095
         Left            =   2520
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
   End
   Begin VB.Menu MenuTitel 
      Caption         =   "&Acties"
      Index           =   0
      Begin VB.Menu Basis 
         Caption         =   "Bedrijf &Openen "
         HelpContextID   =   11000
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu Basis 
         Caption         =   "&Nieuw Bedrijf installeren"
         HelpContextID   =   12000
         Index           =   1
      End
      Begin VB.Menu Basis 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Basis 
         Caption         =   "&Bedrijf Sluiten"
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu Basis 
         Caption         =   "MarSync Starten"
         Index           =   4
      End
      Begin VB.Menu Basis 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu Basis 
         Caption         =   "&Service Pro instellingen"
         Enabled         =   0   'False
         HelpContextID   =   13000
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu Basis 
         Caption         =   "Taakbalk opties vergrendeld"
         Checked         =   -1  'True
         Index           =   7
      End
      Begin VB.Menu Basis 
         Caption         =   "Taakbalk zichtbaar"
         Checked         =   -1  'True
         Index           =   8
      End
      Begin VB.Menu Basis 
         Caption         =   "Manager Standaard 2009"
         Index           =   9
      End
      Begin VB.Menu Basis 
         Caption         =   "Manager Standaard 2005"
         Index           =   10
      End
      Begin VB.Menu Basis 
         Caption         =   "-"
         Index           =   11
         Visible         =   0   'False
      End
      Begin VB.Menu Basis 
         Caption         =   "&XML Rekenbladen"
         Index           =   12
         Shortcut        =   ^X
      End
      Begin VB.Menu Basis 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu Basis 
         Caption         =   "&Afsluiten (Alt+F4)"
         Index           =   14
      End
   End
   Begin VB.Menu MenuTitel 
      Caption         =   "&Systeem"
      Enabled         =   0   'False
      Index           =   1
      Begin VB.Menu Bewerken 
         Caption         =   "&Set-up Boekjaar en Parameters"
         HelpContextID   =   21000
         Index           =   0
      End
      Begin VB.Menu Bewerken 
         Caption         =   "&Datum / periode /boekjaar"
         HelpContextID   =   22000
         Index           =   1
         Shortcut        =   ^D
      End
      Begin VB.Menu Bewerken 
         Caption         =   "&Lay-Out Uitgaand Document"
         HelpContextID   =   23000
         Index           =   2
      End
      Begin VB.Menu Bewerken 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu Bewerken 
         Caption         =   "&Import (Vsoft EDI)"
         HelpContextID   =   24000
         Index           =   4
      End
      Begin VB.Menu Bewerken 
         Caption         =   "S&QL bewerkingen"
         Index           =   5
      End
      Begin VB.Menu Bewerken 
         Caption         =   "&FTP bewerkingen"
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu Bewerken 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu Bewerken 
         Caption         =   "Controle Data"
         Index           =   8
      End
      Begin VB.Menu Bewerken 
         Caption         =   "XML en uw Bedrijf"
         Index           =   9
         Visible         =   0   'False
         Begin VB.Menu mnuXML 
            Caption         =   "ADO XML tester"
            Index           =   0
         End
         Begin VB.Menu mnuXML 
            Caption         =   "XML Data Backup"
            Index           =   1
         End
      End
      Begin VB.Menu Bewerken 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu Bewerken 
         Caption         =   "EURO wisselfunctie"
         Index           =   11
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu MenuTitel 
      Caption         =   "&Venster"
      Index           =   2
      Begin VB.Menu Schikmenu 
         Caption         =   "&Trapsgewijs"
         Index           =   0
      End
      Begin VB.Menu Schikmenu 
         Caption         =   "&Onder elkaar"
         Index           =   1
      End
      Begin VB.Menu Schikmenu 
         Caption         =   "&Naast elkaar"
         Index           =   2
      End
      Begin VB.Menu Schikmenu 
         Caption         =   "&Pictogrammen schikken"
         Index           =   3
      End
      Begin VB.Menu Schikmenu 
         Caption         =   "Open vensters"
         Index           =   4
         WindowList      =   -1  'True
      End
   End
   Begin VB.Menu MenuTitel 
      Caption         =   "&Fiches"
      Enabled         =   0   'False
      HelpContextID   =   39999
      Index           =   3
      Begin VB.Menu Basisbestanden 
         Caption         =   "Dash&board"
         Index           =   0
         Shortcut        =   ^E
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "&Klanten"
         HelpContextID   =   30000
         Index           =   2
         Shortcut        =   ^K
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "&Leveranciers"
         HelpContextID   =   30010
         Index           =   3
         Shortcut        =   ^L
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "&Rekening"
         HelpContextID   =   30020
         Index           =   4
         Shortcut        =   ^R
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "&Artikel/Product/Dienst (OT)"
         HelpContextID   =   30030
         Index           =   5
         Shortcut        =   ^A
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "&Diverse Gebruikersfiches"
         HelpContextID   =   30040
         Index           =   6
         Shortcut        =   ^G
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "Lijs&trapportage"
         Index           =   8
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "&Bestanden Importeren"
         Index           =   9
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "&Grafische voorstellingen"
         Index           =   10
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "-"
         Index           =   11
      End
      Begin VB.Menu Basisbestanden 
         Caption         =   "&Historiek grootboekrekening (scherm)"
         Index           =   13
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu MenuTitel 
      Caption         =   "&Document"
      Enabled         =   0   'False
      Index           =   4
      Begin VB.Menu AV 
         Caption         =   "&Aankoop"
         Index           =   0
         Shortcut        =   +^{F1}
      End
      Begin VB.Menu AV 
         Caption         =   "&Verkoop"
         Index           =   1
         Shortcut        =   +^{F2}
      End
      Begin VB.Menu AV 
         Caption         =   "&Financiëel"
         Index           =   2
         Shortcut        =   +^{F3}
      End
      Begin VB.Menu AV 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu AV 
         Caption         =   "&Kassaverkoop"
         Index           =   4
         Shortcut        =   +^{F6}
      End
      Begin VB.Menu AV 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu AV 
         Caption         =   "&Rekeninguittreksel"
         Index           =   6
      End
      Begin VB.Menu AV 
         Caption         =   "&Electronisch Betalingsverkeer"
         Index           =   7
      End
      Begin VB.Menu AV 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu AV 
         Caption         =   "&StandaardKostPrijsKaart"
         Index           =   9
      End
      Begin VB.Menu AV 
         Caption         =   "&Briefwisseling"
         HelpContextID   =   40900
         Index           =   10
         Shortcut        =   ^B
      End
      Begin VB.Menu AV 
         Caption         =   "Domiciliering Schuldeiser"
         Index           =   11
      End
   End
   Begin VB.Menu MenuTitel 
      Caption         =   "&Boekhouding"
      Enabled         =   0   'False
      Index           =   5
      Begin VB.Menu Boekhouding 
         Caption         =   "&Diverse Posten"
         Index           =   0
         Shortcut        =   +^{F5}
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "&Aankoopboek"
         Index           =   2
         Shortcut        =   +{F1}
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "&Verkoopboek"
         Index           =   3
         Shortcut        =   +{F2}
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "&Financiële boeken"
         Index           =   4
         Shortcut        =   +{F3}
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "&Diverse postenboek"
         Index           =   5
         Shortcut        =   +{F5}
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "BTW-kwartaallisting E.U."
         Enabled         =   0   'False
         Index           =   7
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "BTW-jaarlisting Binnenland"
         HelpContextID   =   50700
         Index           =   8
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "Intrastat 19"
         Index           =   9
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "Intrastat 29"
         Index           =   10
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "Stand BTW-aangifte"
         Index           =   11
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "Kontrole Inventaris Produkten"
         Index           =   13
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "&Historieken Rekeningen"
         Index           =   14
         Shortcut        =   +{F11}
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "Balans &Klanten"
         Index           =   15
         Shortcut        =   +{F7}
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "TopDown Klanten"
         Index           =   16
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "Balans &Leveranciers"
         Index           =   17
         Shortcut        =   +{F8}
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "TopDown Leveranciers"
         Index           =   18
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "&Proef- en Saldibalans"
         Index           =   19
         Shortcut        =   +{F12}
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "-"
         Index           =   20
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "&Eindrapportage"
         Index           =   21
      End
      Begin VB.Menu Boekhouding 
         Caption         =   "&Overgangsprogramma's"
         Index           =   22
         Begin VB.Menu Overgang 
            Caption         =   "&Nieuw boekjaar"
            Index           =   0
         End
         Begin VB.Menu Overgang 
            Caption         =   "&Opkuis bestanden"
            Index           =   1
         End
      End
   End
   Begin VB.Menu MenuTitel 
      Caption         =   "&Contractbeheer"
      Enabled         =   0   'False
      Index           =   6
      Begin VB.Menu pBeheer 
         Caption         =   "&Opmaak kwijtingen Post"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu pBeheer 
         Caption         =   "&Inboeken kwijting(en)"
         Index           =   2
         Shortcut        =   ^I
      End
      Begin VB.Menu pBeheer 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu pBeheer 
         Caption         =   "&Vrije Bericht AS1 (020) of post"
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu pBeheer 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu pBeheer 
         Caption         =   "Toepassingen maatschappijen opvragen"
         Index           =   6
      End
      Begin VB.Menu pBeheer 
         Caption         =   "Toepassingen maatschappijen kontroleren"
         Index           =   7
      End
      Begin VB.Menu pBeheer 
         Caption         =   "AS2 en GBO Uitwisseling"
         Index           =   8
         Shortcut        =   ^T
      End
      Begin VB.Menu pBeheer 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu pBeheer 
         Caption         =   "Groene Kaart Auto"
         Index           =   10
      End
      Begin VB.Menu pBeheer 
         Caption         =   "Domicilieringen"
         Index           =   11
      End
   End
   Begin VB.Menu MenuTitel 
      Caption         =   "Cloud4MAR"
      Index           =   7
      Begin VB.Menu DNN 
         Caption         =   "Manuele Documenten (serverlocatie)"
         Index           =   0
      End
      Begin VB.Menu DNN 
         Caption         =   "Archief (serverlocatie)"
         Index           =   1
      End
      Begin VB.Menu DNN 
         Caption         =   "Instellingen"
         Index           =   2
      End
      Begin VB.Menu DNN 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu DNN 
         Caption         =   "https://vsoft.be"
         Index           =   4
      End
   End
   Begin VB.Menu MenuTitel 
      Caption         =   "&Instellingen"
      Index           =   8
      Begin VB.Menu Instellingen 
         Caption         =   "&Printer"
         Index           =   0
      End
      Begin VB.Menu Instellingen 
         Caption         =   "&Woordenboek Fiches"
         Index           =   1
      End
      Begin VB.Menu Instellingen 
         Caption         =   "&Database Editor"
         Index           =   2
      End
      Begin VB.Menu Instellingen 
         Caption         =   "&VeldDefinitie"
         Index           =   3
      End
      Begin VB.Menu Instellingen 
         Caption         =   "SQL Server gebruiken"
         Index           =   4
      End
      Begin VB.Menu Instellingen 
         Caption         =   "marnt.arc uitvergroten"
         Enabled         =   0   'False
         Index           =   5
      End
   End
   Begin VB.Menu MenuTitel 
      Caption         =   "&?"
      Index           =   9
      Begin VB.Menu Info 
         Caption         =   "Plesk Hosting"
         Index           =   0
         Begin VB.Menu Hosting 
            Caption         =   "Obsidian v18.0.65"
            Index           =   0
         End
         Begin VB.Menu Hosting 
            Caption         =   "Webmail rv-services"
            Index           =   1
         End
      End
      Begin VB.Menu Info 
         Caption         =   "Commissie Boekhoudkundige Normen"
         Index           =   1
      End
      Begin VB.Menu Info 
         Caption         =   "HTML Help marIntegraal &OT"
         Enabled         =   0   'False
         Index           =   2
      End
      Begin VB.Menu Info 
         Caption         =   "&Licentie toewijzing"
         Index           =   3
      End
      Begin VB.Menu Info 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu Info 
         Caption         =   "Command Prompt"
         Index           =   5
      End
      Begin VB.Menu Info 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu Info 
         Caption         =   "Peppol"
         Index           =   7
         Begin VB.Menu Peppol 
            Caption         =   "Document Validator"
            Index           =   0
            Shortcut        =   ^U
         End
         Begin VB.Menu Peppol 
            Caption         =   "UBL.BE Docs Billing 3.0"
            Index           =   1
         End
         Begin VB.Menu Peppol 
            Caption         =   "ClickOnce MarSync"
            Index           =   2
         End
      End
   End
End
Attribute VB_Name = "Mim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Sub InitEerst()

FULL_LINE = String$(128, 173)

TABLEDEF_ONT(TABLE_VARIOUS) = "0000000.ONT"         '00
TABLEDEF_ONT(TABLE_CUSTOMERS) = "0010000.ONT"            '01
TABLEDEF_ONT(TABLE_SUPPLIERS) = "0020000.ONT"      '02
TABLEDEF_ONT(TABLE_LEDGERACCOUNTS) = "0030000.ONT"         '03
TABLEDEF_ONT(TABLE_PRODUCTS) = "0040000.ONT"          '04
TABLEDEF_ONT(TABLE_JOURNAL) = "0600000.ONT"         '05
TABLEDEF_ONT(TABLE_INVOICES) = "0200000.ONT"         '06
TABLEDEF_ONT(TABLE_CONTRACTS) = "0700000.ONT"            '07
TABLEDEF_ONT(TABLE_DUMMY) = "90DUMMY.ONT"            '08
TABLEDEF_ONT(TABLE_COUNTERS) = "00.ONT"                '09

bstNaam(TABLE_VARIOUS) = "Allerlei"        '00
bstNaam(TABLE_CUSTOMERS) = "Klanten"            '01
bstNaam(TABLE_SUPPLIERS) = "Leveranciers" '02
bstNaam(TABLE_LEDGERACCOUNTS) = "Rekeningen"      '03
bstNaam(TABLE_PRODUCTS) = "Produkten"        '04
bstNaam(TABLE_JOURNAL) = "Journalen"       '05
bstNaam(TABLE_INVOICES) = "dokumenten"      '06
bstNaam(TABLE_CONTRACTS) = "Polissen"           '07
bstNaam(TABLE_DUMMY) = "TmpBestand"         '08
bstNaam(TABLE_COUNTERS) = "Tell"              '09

DAYS_IN_MONTH(1) = 31
DAYS_IN_MONTH(2) = 29
DAYS_IN_MONTH(3) = 31
DAYS_IN_MONTH(4) = 30
DAYS_IN_MONTH(5) = 31
DAYS_IN_MONTH(6) = 30
DAYS_IN_MONTH(7) = 31
DAYS_IN_MONTH(8) = 31
DAYS_IN_MONTH(9) = 30
DAYS_IN_MONTH(10) = 31
DAYS_IN_MONTH(11) = 30
DAYS_IN_MONTH(12) = 31

MONTH_AS_TEXT(1) = "Januari  "
MONTH_AS_TEXT(2) = "Februari "
MONTH_AS_TEXT(3) = "Maart    "
MONTH_AS_TEXT(4) = "April    "
MONTH_AS_TEXT(5) = "Mei      "
MONTH_AS_TEXT(6) = "Juni     "
MONTH_AS_TEXT(7) = "Juli     "
MONTH_AS_TEXT(8) = "Augustus "
MONTH_AS_TEXT(9) = "September"
MONTH_AS_TEXT(10) = "October  "
MONTH_AS_TEXT(11) = "November "
MONTH_AS_TEXT(12) = "December "

End Sub

Sub InstalleerAllePrinters()

Dim PrinterTekst As String
Dim FlTemp As Integer

On Error Resume Next
CmbDokumentPrinter.Clear
CmbLijstPrinter.Clear
cmbKassaPrinter.Clear
If Printers.Count = 0 Then MsgBox "Installeer eerst een printerdriver a.u.b. (ook al beschikt U nog over geen printer.  Bij blijvende problemen in netwerk gelieve op te starten met het command '/P=FALSE' en uw netwerkbeheerder te informeren)", vbExclamation: End
For COUNT_TO = 0 To Printers.Count - 1
    PrinterTekst = Printers(COUNT_TO).Port + " " + Printers(COUNT_TO).DeviceName
    CmbDokumentPrinter.AddItem PrinterTekst
    CmbLijstPrinter.AddItem PrinterTekst
    cmbKassaPrinter.AddItem PrinterTekst
Next
PrintersTweedePoging:
If Dir(PROGRAM_LOCATION + "Def\prnt2000.SET") <> "" Then
    Dim dokumentPrinterNaam As String
    Dim LijstPrinterNaam As String
    Dim KassaPrinterNaam As String
    FlTemp = FreeFile
    Open PROGRAM_LOCATION + "Def\prnt2000.SET" For Input As FlTemp
    Err = 0
    Line Input #FlTemp, dokumentPrinterNaam
    Line Input #FlTemp, LijstPrinterNaam
    Line Input #FlTemp, KassaPrinterNaam
    Close FlTemp
    If Err Then
        Kill PROGRAM_LOCATION + "Def\prnt2000.SET"
        GoTo PrintersTweedePoging
    End If
    For COUNT_TO = 0 To Printers.Count - 1
        If dokumentPrinterNaam = Printers(COUNT_TO).Port + " " + Printers(COUNT_TO).DeviceName Then CmbDokumentPrinter.ListIndex = COUNT_TO
        If LijstPrinterNaam = Printers(COUNT_TO).Port + " " + Printers(COUNT_TO).DeviceName Then CmbLijstPrinter.ListIndex = COUNT_TO
        If KassaPrinterNaam = Printers(COUNT_TO).Port + " " + Printers(COUNT_TO).DeviceName Then cmbKassaPrinter.ListIndex = COUNT_TO
    Next
Else
    FlTemp = FreeFile
    Open PROGRAM_LOCATION + "Def\prnt2000.SET" For Output As FlTemp
    Print #FlTemp, Printers(0).Port + " " + Printers(0).DeviceName
    Print #FlTemp, Printers(0).Port + " " + Printers(0).DeviceName
    Print #FlTemp, Printers(0).Port + " " + Printers(0).DeviceName
    Close FlTemp
    MsgBox "prnt2000.SET wordt automatisch ingesteld op de eerste printer en eerste papierlade voor afdruk van dokumenten, lijsten en kassatiketten.  Pas zelf aan indien nodig"
    BeWaarTekst App.Title, "dokumentPRINTER", "01"
    BeWaarTekst App.Title, "LIJSTPRINTER", "01"
    BeWaarTekst App.Title, "KASSAPRINTER", "01"
    GoTo PrintersTweedePoging
End If

End Sub

Function KtrlTVSOMERS()

Msg = Msg

Dim rsJourHier As ADODB.Recordset
Dim rsdokumentHier As ADODB.Recordset

Dim TotaalTeBetalen As Double
Dim TeWijzigen As Double

Dim COUNT_TO As Integer

Set rsJourHier = New ADODB.Recordset
Set rsdokumentHier = New ADODB.Recordset

Nogeens:
Msg = InputBox("Geef het te kontroleren Nummer (bvb. 488000)", , "488000")
If Msg = "" Then Exit Function

rsJourHier.CursorLocation = adUseClient
rsdokumentHier.CursorLocation = adUseClient
rsJourHier.Open "SELECT * FROM Journalen WHERE v019 = '" + vSet(Msg, 7) + "'", adntDB, adOpenForwardOnly, adLockOptimistic
If rsJourHier.RecordCount Then
    MsgBox "Kontrole voor " & rsJourHier.RecordCount & " records"
    rsJourHier.MoveFirst
    Do While Not rsJourHier.EOF
        If Left(RV(rsJourHier, "v033"), 1) = "V" And RV(rsJourHier, "v038") = "" Then
            rsdokumentHier.Open "SELECT * FROM dokumenten WHERE v033 = '" + rsJourHier("v033") + "'", adntDB, adOpenDynamic, adLockOptimistic
            If rsdokumentHier.RecordCount Then
                ' alles optellen en indien niet overeenkomend dan
                ' wijzigen
                TotaalTeBetalen = 0
                For COUNT_TO = 55 To 64
                    TotaalTeBetalen = TotaalTeBetalen + Val(RV(rsdokumentHier, "v0" + Format(COUNT_TO)))
                Next
                If Val(RV(rsdokumentHier, "v037")) = TotaalTeBetalen Then
                ElseIf Val(RV(rsdokumentHier, "v037")) > TotaalTeBetalen Then
                    If Left(RV(rsJourHier, "v033"), 2) = "V0" Then
                        'Faktuur
                        TeWijzigen = Val(RV(rsdokumentHier, "v055")) - Val(RV(rsJourHier, "v068"))
                        KtrlBox = MsgBox("Faktuur " + rsJourHier("v033") + ", vak 55: " + RV(rsdokumentHier, "v055") + " verhogen tot " + Format(TeWijzigen), vbQuestion + vbDefaultButton1 + vbYesNo)
                        If KtrlBox = vbYes Then
                            rsdokumentHier("v055") = Format(TeWijzigen)
                            rsdokumentHier.Update
                        End If
                    Else
                        'Creditnota
                        TeWijzigen = Val(RV(rsdokumentHier, "v063")) + Val(RV(rsJourHier, "v068"))
                        KtrlBox = MsgBox("Faktuur " + rsJourHier("v033") + ", vak 63: " + RV(rsdokumentHier, "v063") + " verhogen tot " + Format(TeWijzigen), vbQuestion + vbDefaultButton1 + vbYesNo)
                        If KtrlBox = vbYes Then
                            rsdokumentHier("v063") = Format(TeWijzigen)
                            rsdokumentHier.Update
                        End If
                    End If
                End If
            End If
            rsdokumentHier.Close
        End If
        rsJourHier.MoveNext
    Loop
    rsJourHier.Close
Else
    rsJourHier.Close
    GoTo Nogeens
End If
Set rsJourHier = Nothing
Set rsdokumentHier = Nothing

End Function

Private Sub AV_Click(Index As Integer)

On Error Resume Next
Select Case Index
    Case 0
        If ACTIVE_BOOKYEAR > 1 Then MsgBox "Boekhoudkundige bewerkingen enkel mogelijk in hoogste 2 boekjaren !": Exit Sub
        DirekteAankoop.Show

    Case 1
        If ACTIVE_BOOKYEAR > 1 Then MsgBox "Boekhoudkundige bewerkingen enkel mogelijk in hoogste 2 boekjaren !": Exit Sub
        DirekteVerkoop.Show
    
    Case 2
        If ACTIVE_BOOKYEAR > 1 Then MsgBox "Boekhoudkundige bewerkingen enkel mogelijk in hoogste 2 boekjaren !": Exit Sub
        InbrengFinancieel.Show
            
    Case 4
        frmKassaVerkoop.Show 1
                        
    Case 6
        Rappelklanten.Show 1
    Case 7
        frmOGM.Show
        
    Case 9
        If ACTIVE_BOOKYEAR > 0 Then MsgBox "Opmaak uitsluiten in hoogste boekjaar mogelijk !": Exit Sub
        StandaardKostPrijsKaart.Show
        
    Case 10
        VrijBericht.Show
        
    Case 11
        
        Dim r As Long
        If ShellExecuteWithFallback("https://cdd.vsoft.be") = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
        
End Select


End Sub

Private Sub Basis_Click(Index As Integer)
Dim sVFL        As Integer
Dim checkLock As String

On Error Resume Next
Msg = "Hierna worden eerst alle bestanden en openstaande vensters van een actief bedrijf gesloten."

Select Case Index
    Case 0
        If GaVerder(Msg, 1, "Bedrijf Openen") Then
            Ktrl = 100
            AutoUnloadBedrijf
            On Local Error Resume Next
            ntDB.Close
            On Local Error GoTo 0
            BedrijfOpenen.Show
        End If
    
    Case 1
        If GaVerder(Msg, 1, "Nieuw Bedrijf maken") Then
            Ktrl = 100
            AutoUnloadBedrijf
            'MsgBox "Tijdelijk niet mogelijk in MSDE overgangsfase.  Gebruik versie 6.1.80 of lager voor aanmaak nieuw bedrijf a.u.b.", vbInformation
            'Exit Sub
            NieuwBedrijf.Show 1
        End If
        
    Case 3
        checkLock = LOCATION_COMPANYDATA
        
        Ktrl = 100
        AutoUnloadBedrijf
        On Local Error Resume Next
        ntDB.Close
        On Local Error GoTo 0
        DoEvents
        If Dir(checkLock & "marnt.ldb") <> "" Then
            MsgBox "Er zijn nog toestellen actief met dit bedrijf. Gelieve deze zeker te sluiten vooralaar MarSync te gebruiken", vbInformation
        End If
        
    Case 4
        checkLock = LOCATION_COMPANYDATA
        
        If checkLock = "" Then
        Else
            Ktrl = 100
            AutoUnloadBedrijf
            On Local Error Resume Next
            ntDB.Close
            On Local Error GoTo 0
            DoEvents
        End If
        Call DetectClickOnceShortcut
        
    Case 6
        
        If Dir$(LOCATION_COMPANYDATA + "Menu.Def") = "" Then Exit Sub
        Msg = LOCATION_COMPANYDATA + "Menu.Def voor 'Light' versie vernietigen !  Bent U zeker ?"
        Ktrl = MsgBox(Msg, 292)
        If Ktrl = 6 Then
            Kill LOCATION_COMPANYDATA + "Menu.Def"
        End If
        
    Case 7
        Basis(7).Checked = Not Basis(7).Checked
        If Basis(7).Checked Then
            Me.cbMovie1.Enabled = False
            Me.cbMovie2.Enabled = False
            Me.cbMovie3.Enabled = False
                        
            CmbDokumentPrinter.Enabled = False
            CmbLijstPrinter.Enabled = False
            cmbKassaPrinter.Enabled = False
            cmbpaperbin.Visible = False
            'PrinterFont.Visible = False
        Else
            Me.cbMovie1.Enabled = True
            Me.cbMovie2.Enabled = True
            Me.cbMovie3.Enabled = True
            
            CmbDokumentPrinter.Enabled = True
            CmbLijstPrinter.Enabled = True
            cmbKassaPrinter.Enabled = True
            cmdWegBoekModus.Enabled = True
            cmbpaperbin.Clear
            cmbpaperbin.AddItem "01: Papieraanvoer vanuit bovenste lade (default)"
            cmbpaperbin.AddItem "02: Papieraanvoer vanuit onderste lade"
            cmbpaperbin.AddItem "03: Papieraanvoer vanuit middenste lade"
            cmbpaperbin.AddItem "04: Wacht voor manuele invoeging van elk blad papier"
            cmbpaperbin.AddItem "05: Gebruik omslagen via enveloppe lade"
            cmbpaperbin.AddItem "06: Enveloppe lade gebruiken met manuele invoer"
            cmbpaperbin.AddItem "07: Gebruik de standaard invoerlade (Default2)"
            cmbpaperbin.AddItem "08: Gebruik papier aangestuurd door de tractor"
            cmbpaperbin.AddItem "09: Papieraanvoer via de kleine lade"
            cmbpaperbin.AddItem "10: Papieraanvoer via de grote lade"
            'cmbpaperbin.AddItem "14: Gebruik papier vanuit de gekoppelde cassette"
            cmbpaperbin.Visible = True
            cmbpaperbin.ListIndex = 0
            'PrinterFont.Visible = True
        End If
        
    Case 8
        Basis(8).Checked = Not Basis(8).Checked
        If Basis(8).Checked Then
            tbToolBar.Visible = True
        Else
            tbToolBar.Visible = False
        End If
        BeWaarTekst App.Title, "TaakBalkZichtbaar", (tbToolBar.Visible)
            
    Case 9
        '1152 x 864
        Mim.WindowState = 0
        Mim.Height = 12400
        Mim.Width = 17300
        Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
        
    Case 10
        Mim.WindowState = 0
        Mim.Height = 10410
        Mim.Width = 14310
        Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
                
    'case
    'mim.windowstate=0
        Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
                
    Case 12
        'frmSimpleScript.Show
        frmRekenBlad.Show
                                       
    Case 14
        Ktrl = 100
        Unload Mim
End Select

End Sub

Private Sub BasisBestanden_Click(Index As Integer)

Select Case Index
    Case 0
        Venster.Show
        
    Case 2 To 4
        BasisB(Index - 1).WindowState = 0
        BasisB(Index - 1).SetFocus
    
    Case 5
        'BasisB(Index + 1).WindowState = 0
        'BasisB(Index + 1).SetFocus
        frmProduktFiche.Show
    Case 6
        xDokument.WindowState = 0
        xDokument.SetFocus
    Case 8
        SQLLijsten.Show
    Case 9
        Importeren.Show
    Case 10
        'Form1.Show
        'SnelHelpPrint "In ontwikkeling " & Now, BL_LOGGING
    Case 13
        HistoriekSQL.Show
End Select

End Sub

Private Sub Bewerken_Click(Index As Integer)
Dim ImportBestand As String
Dim X As Integer

Select Case Index
    Case 0
        SetupEnParameters.Show
    Case 1
        BJPERDAT.WindowState = 0
        BJPERDAT.PeriodeBoekjaar.SetFocus
    Case 2
        LayOutpdfDokument.Show
        
    Case 4
        If VsoftLog Then MsgBox "U draait Light modus en wil tevens importeren !!  Kan en mag niet...": Exit Sub

        On Local Error GoTo ErrorJump

        Mim.Teken.filename = "NTImport.Log"
        Mim.Teken.Filter = "Enkel (NTImport.Log)|NTImport.Log"
        Mim.Teken.CancelError = True
        Mim.Teken.Action = 1
        ImportBestand = Mim.Teken.filename + " "
        If Dir$(ImportBestand) = "" Then Exit Sub

        On Local Error GoTo 0

              Msg = "Gegevens uit " + ImportBestand + " invoegen in actief bedrijf." + vbCrLf + vbCrLf
        Msg = Msg + "Bent U zeker ?"
        Ktrl = MsgBox(Msg, 292)
        If Ktrl = 6 Then
            BlokImport ImportBestand
        Else
            Exit Sub
        End If
    
    Case 5
        On Error Resume Next
        frmADOSQLOpzoeker.Show
            
    Case 6
        'frmMarNtInetUpgrade.Show
        
    Case 8
        X = KontroleerDatabase()
        
    Case 11
        'Wissel.Show
               
End Select

ErrorJump:
Exit Sub
Resume

End Sub

Private Sub BlokImport(BestandEnPath As String)

Dim flHier As Integer
Dim RecordHier As String
Dim AktieHier As String
Dim RecIndex As Integer
Dim BTRVType As String

MousePointer = 11
bClose 99
B_MODUS = -1
flHier = FreeFile
Open BestandEnPath For Input As flHier
Dim X As Integer
If BAModus = 1 Then
    BeginTrans
End If
Mim.WindowState = 0
Mim.Enabled = False

Do While Not EOF(flHier)
    XDoEvents = DoEvents()
    Line Input #flHier, RecordHier
    Select Case Mid(RecordHier, 8, 8)
        
        Case "ALLERLEI" '0
            Fl = TABLE_VARIOUS
        Case "KLANTEN " '1
            Fl = TABLE_CUSTOMERS
        Case "LEVERANC" '2
            Fl = TABLE_SUPPLIERS
        Case "REKENING" '3
            Fl = TABLE_LEDGERACCOUNTS
        Case "PRODUKT " '4
            Fl = TABLE_PRODUCTS
        Case "dokument" '5
            Fl = TABLE_INVOICES
        Case "JOURNAAL" '6
            Fl = TABLE_JOURNAL
        Case "POLISSEN" '7
            Fl = TABLE_CONTRACTS
        Case "TELLERS " '8
            Fl = TABLE_COUNTERS
        Case Else
            MsgBox "stop"
    End Select

    BTRVType = Mid(RecordHier, 1, 6)
    RecIndex = Val(Mid(RecordHier, 7, 1))
    TLB_RECORD(Fl) = Mid(RecordHier, 16)
    Select Case BTRVType
        Case "INSERT", "UPDATE", "GET   ", "GETOG "
            Select Case BTRVType
                Case "INSERT", "UPDATE"
                    'If Fl = TABLE_CONTRACTS Then
                    '    vBib Fl, vSet(vBibTekst(TABLE_CONTRACTS, "#v164 #"), 2) + vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12) + vSet(vBibTekst(TABLE_CONTRACTS, "#A010 #"), 4) + vSet(vBibTekst(TABLE_CONTRACTS, "#A000 #"), 12), "v167" 'MaandKlantMaatschappijPolis
                    'End If
                    Select Case Fl
                        Case TABLE_CUSTOMERS, TABLE_SUPPLIERS
                            If vBibTekst(Fl, "#A10C #") = "" Then vBib Fl, "2", "A10C"   'Taalkode
                            If vBibTekst(Fl, "#v149 #") = "" Then vBib Fl, "002", "v149" 'Landnummer  ISO kode
                            If vBibTekst(Fl, "#A109 #") = "" Then vBib Fl, "B  ", "A109" 'Landkode Postkantoor
                            If vBibTekst(Fl, "#v150 #") = "" Then vBib Fl, "BE", "v150" 'Landkode    ISO kode
                            If vBibTekst(Fl, "#vs03 #") = "" Then vBib Fl, "BEF", "vs03" 'Munteenheid ISO kode
                            If vBibTekst(Fl, "#vs07 #") = "" Then vBib Fl, "1", "vs07" 'exemplaren dokumenten
                        Case TABLE_LEDGERACCOUNTS
                            If vBibTekst(Fl, "#v032 #") = "" Then vBib Fl, "O", "v032" 'Budgetcode

                    End Select
                    SnelHelpPrint FVT(Fl, RecIndex), BL_LOGGING
                    Select Case BTRVType
                        Case "INSERT"
                            bInsert Fl, RecIndex
                        Case "UPDATE"
                            bUpdate Fl, RecIndex
                        Case Else
                            MsgBox "Logikafout !"
                    End Select
                
                Case "GET   ", "GETOG "
                    Select Case BTRVType
                        Case "GET   "
                            bGet Fl, RecIndex, TLB_RECORD(Fl)
                        Case "GETOG "
                            bGetOrGreater Fl, RecIndex, TLB_RECORD(Fl)
                        Case Else
                            MsgBox "Logikafout !"
                    End Select
                Case Else
                    MsgBox "Logikafout !"
            End Select
        
        Case "FIRST "
            bFirst Fl, RecIndex
        Case "LAST  "
            bLast Fl, RecIndex
        Case "DELETE"
            bDelete Fl
        Case "NEXT  "
            bNext Fl
        Case "PREV  "
            bPrev Fl
        Case Else
            MsgBox "LogikaFout !"
    End Select
Loop
Close flHier
bClose 99
If BAModus = 1 Then
    CommitTrans
End If
MousePointer = 0
Kill BestandEnPath
MsgBox "Bijwerking met succes beëindigd !"
Mim.WindowState = 2
Mim.Enabled = True

End Sub

Private Sub Boekhouding_Click(Index As Integer)

On Error Resume Next
Select Case Index
    Case 0
        If ACTIVE_BOOKYEAR > 1 Then MsgBox "Boekhoudkundige bewerkingen enkel mogelijk in hoogste 2 boekjaren !": Exit Sub
        DiversePosten.Show
    
    Case 2
        aIndex = TABLE_SUPPLIERS
        AVBoek.Show 1

    Case 3
        aIndex = TABLE_CUSTOMERS
        AVBoek.Show 1

    Case 4
        FinancieelBoek.Show 1
        
    Case 5
        DiversePostenBoek.Show 1

    Case 8
        JaarBtwBelgie.Show 1

    Case 9 'Intrastat 19
        aIndex = 19
        LijstIntrastat19.Show 1
    Case 10
        aIndex = 29
        LijstIntrastat19.Show 1
    
    Case 11
        frmBTWAangifte.Show 1
        
    Case 13
        If ACTIVE_BOOKYEAR <> 0 Then MsgBox "Stock Inventarisatie enkel mogelijk in hoogste boekjaar !": Exit Sub
        InventarisOpgave.Show 1

    Case 14
        HistoriekPrint.Show 1
        
    Case 15
        aIndex = TABLE_CUSTOMERS
        BalansKL.Show 1
        
    Case 16
        aIndex = TABLE_CUSTOMERS
        TopDown.Show 1

    Case 17
        aIndex = TABLE_SUPPLIERS
        BalansKL.Show 1

    Case 18
        aIndex = TABLE_SUPPLIERS
        TopDown.Show 1

    Case 19
        ProefEnSaldi.Show 1

    Case 21
        Rapportage.Show 1
       
End Select
If Index <> 0 Then
    Mim.SetFocus
End If

End Sub

Private Function GaVerder(Bericht As String, BedrijfOpenKontrole As Integer, Titel As String) As Integer

If BedrijfOpenKontrole Then
    If BasisB(1).Enabled = True Then
    Else
        GaVerder = True
        Exit Function
    End If
End If

KtrlBox = MsgBox(Bericht + vbCrLf + vbCrLf + "Bent U zeker ?", 292, Titel)
If KtrlBox = 6 Then
    GaVerder = True
Else
    GaVerder = False
End If

End Function

Private Sub cbMovie1_Click()

    Dim ZoekHTM As Variant
    Dim r As Long
    
    ZoekHTM = App.path & "\marMovies\InternetDownload.html"
    If Dir(ZoekHTM) <> "" Then
        Err = 0
        On Error Resume Next
        If Err Then MsgBox Error: Exit Sub
        If ShellExecuteWithFallback(ZoekHTM) = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
        DoEvents
    End If

End Sub

Private Sub cbMovie2_Click()

    Dim ZoekHTM As Variant
    Dim r As Long
    
    ZoekHTM = App.path & "\marMovies\KoppelKostenRekening.html"
    If Dir(ZoekHTM) <> "" Then
        Err = 0
        On Error Resume Next
        If Err Then MsgBox Error: Exit Sub
        If ShellExecuteWithFallback(ZoekHTM) = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
        DoEvents
    End If
End Sub

Private Sub cbMovie3_Click()

    Dim ZoekHTM As Variant
    Dim r As Long
    
    ZoekHTM = App.path & "\marMovies\btwIntervat.html"
    If Dir(ZoekHTM) <> "" Then
        Err = 0
        On Error Resume Next
        If Err Then MsgBox Error: Exit Sub
        If ShellExecuteWithFallback(ZoekHTM) = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
        DoEvents
    End If

End Sub

Private Sub CmbDokumentPrinter_Click()

    Dim dokumentPrinterNaam As String
    Dim LijstPrinterNaam As String
    Dim KassaPrinterNaam As String
    
    Dim FlTemp As Integer
    
    FlTemp = FreeFile
    Open PROGRAM_LOCATION + "Def\prnt2000.SET" For Output As FlTemp
    Print #FlTemp, CmbDokumentPrinter.text
    Print #FlTemp, CmbLijstPrinter.text
    Print #FlTemp, cmbKassaPrinter.text
    Close FlTemp
    dokumentPrinterNr = CmbDokumentPrinter.ListIndex
    
End Sub

Private Sub CmbdokumentPrinter_Validate(Cancel As Boolean)

Msg = "Voor dokumentPRINTER papierinvoer :" & vbLf & vbLf
Msg = Msg & cmbpaperbin.text & vbLf & vbLf
Msg = Msg & "Bent U zeker"
KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton1)
If KtrlBox = vbYes Then
    On Error Resume Next
    Set Printer = Printers(dokumentPrinterNr)
    Printer.PaperBin = Left(cmbpaperbin.text, 2)
    If Err Then
        MsgBox Error
        Exit Sub
    Else
        BeWaarTekst App.Title, "dokumentPRINTER", Left(cmbpaperbin.text, 2)
    End If
End If

End Sub

Private Sub cmbKassaPrinter_Click()

    Dim dokumentPrinterNaam As String
    Dim LijstPrinterNaam As String
    Dim KassaPrinterNaam As String
    
    Dim FlTemp As Integer
    FlTemp = FreeFile
    Open PROGRAM_LOCATION + "Def\prnt2000.SET" For Output As FlTemp
    Print #FlTemp, CmbDokumentPrinter.text
    Print #FlTemp, CmbLijstPrinter.text
    Print #FlTemp, cmbKassaPrinter.text
    Close FlTemp
    KassaPrinterNr = cmbKassaPrinter.ListIndex
    
End Sub

Private Sub cmbKassaPrinter_Validate(Cancel As Boolean)

Msg = "Voor KASSAPRINTER papierinvoer :" & vbLf & vbLf
Msg = Msg & cmbpaperbin.text & vbLf & vbLf
Msg = Msg & "Bent U zeker"
KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton1)
If KtrlBox = vbYes Then
    On Error Resume Next
    Set Printer = Printers(KassaPrinterNr)
    Printer.PaperBin = Left(cmbpaperbin.text, 2)
    If Err Then
        MsgBox Error
        Exit Sub
    Else
        BeWaarTekst App.Title, "KASSAPRINTER", Left(cmbpaperbin.text, 2)
    End If
End If

End Sub

Private Sub CmbLijstPrinter_Click()

    Dim dokumentPrinterNaam As String
    Dim LijstPrinterNaam As String
    Dim KassaPrinterNaam As String
    
    Dim FlTemp As Integer
    FlTemp = FreeFile
    Open PROGRAM_LOCATION + "Def\prnt2000.SET" For Output As FlTemp
    Print #FlTemp, CmbDokumentPrinter.text
    Print #FlTemp, CmbLijstPrinter.text
    Print #FlTemp, cmbKassaPrinter.text
    Close FlTemp
    LijstPrinterNr = CmbLijstPrinter.ListIndex
    
End Sub


Private Sub CmbLijstPrinter_Validate(Cancel As Boolean)

Msg = "Voor LIJSTPRINTER papierinvoer :" & vbLf & vbLf
Msg = Msg & cmbpaperbin.text & vbLf & vbLf
Msg = Msg & "Bent U zeker"
KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton1)
If KtrlBox = vbYes Then
    On Error Resume Next
    Set Printer = Printers(LijstPrinterNr)
    Printer.PaperBin = Left(cmbpaperbin.text, 2)
    If Err Then
        MsgBox Error
        Exit Sub
    Else
        BeWaarTekst App.Title, "LIJSTPRINTER", Left(cmbpaperbin.text, 2)
    End If
End If

End Sub


Private Sub cmdWegBoekModus_Click()

BeWaarTekst "Algemeen", "BoekInfoModus", cmdWegBoekModus.text

End Sub


Private Sub cmdWegBoekModus_GotFocus()

   'MsgBox "Maak uw keuze en druk daarna ENTER om te vergrendelen.  Kies via Acties/Menubalk Opties om te ontgrendelen en andere keuze te maken", vbInformation
   
End Sub

Private Sub cmdWegBoekModus_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then cmdWegBoekModus.Enabled = False

End Sub

Private Sub DatumVerwerking_Change()

    MIM_GLOBAL_DATE = Format(DatumVerwerking.Value, "dd/mm/yyyy")
    BJPERDAT.DatumVerwerking.Value = DatumVerwerking.Value
            
End Sub

Private Sub dnn_Click(Index As Integer)

'Dim frmB As New frmBrowser
Dim r As Long
Dim strAdress As String

On Local Error Resume Next
Err = 0
Select Case Index
    Case 0
    
        If LaadTekst("dnnInstellingen", "Mario") = "" Then
            MsgBox "Eerst uw URL instellen a.u.b. probeer later opnieuw", vbExclamation
            frmDNNinstellingen.Show 1
            Exit Sub
        End If
        
        strAdress = LaadTekst("dnnInstellingen", "Mario")
        If ShellExecuteWithFallback(strAdress) = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
        
    Case 1
           
        If Err Then MsgBox Error: Exit Sub
        
        If LaadTekst("dnnInstellingen", "Archief") = "" Then
            MsgBox "Eerst uw URL instellen a.u.b. probeer later opnieuw", vbExclamation
            frmDNNinstellingen.Show 1
            Exit Sub
        End If
        
        strAdress = LaadTekst("dnnInstellingen", "Archief")
        If ShellExecuteWithFallback(strAdress) = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
                
    Case 2
        frmDNNinstellingen.Show 1
        
    Case 4
            
        If Err Then MsgBox Error: Exit Sub
        If ShellExecuteWithFallback("https://rv.be/accounting") = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
        
Case Else

End Select

End Sub

Private Sub FontGrootte_LostFocus()

Printer.FontSize = FontGrootte.text
FontGrootte.text = Printer.FontSize

End Sub

Private Sub Hosting_Click(Index As Integer)

'Dim frmB As New frmBrowser
Dim dKtrl As Double
Dim path As String * 260
Dim deFolder As String
Dim deSubFolder As String

Dim ret As Integer
Dim r As Long

On Local Error Resume Next
Err = 0
Select Case Index
   Case 0
        If Err Then MsgBox Error: Exit Sub
        If ShellExecuteWithFallback("https://web24.foxxl.com:8443") = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
    
    Case 1
        If Err Then MsgBox Error: Exit Sub
        If ShellExecuteWithFallback("https://webmail.rv-services.be") = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
End Select

End Sub

Private Sub Info_Click(Index As Integer)

'Dim frmB As New frmBrowser
Dim dKtrl As Double
Dim path As String * 260
Dim deFolder As String
Dim deSubFolder As String

Dim ret As Integer
Dim r As Long

On Local Error Resume Next
Err = 0
Select Case Index
    'Case 0

    Case 1
        If Err Then MsgBox Error: Exit Sub
        'Commissie boekhoudkundige normen
        If ShellExecuteWithFallback("https://www.cbn-cnc.be/nl") = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
            
    Case 2
        If ShellExecuteWithFallback("hh.exe " + PROGRAM_LOCATION + "marot.chm") = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
        'Oude DOS handleiding
        If Err.Number = 53 Then
            MsgBox "U kan de meest recente handleiding in HTML formaat ONLINE op het Internet raadplegen alsook OPHALEN.  Raadpleeg uw Microsoft Certified Internet Dealer of kontakteer ons voor marIntegraal Internet Servicecontract."
        End If
        
    Case 3
        MousePointer = vbHourglass
        'Licentievenster
        Vsoft.Show
        MousePointer = vbNormal
        
    Case 5
        If Err Then MsgBox Error: Exit Sub
        'Command Prompt
        If ShellExecuteWithFallback("cmd") = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
        
    'Case 7
        'QR.Show
        'ShowAndSaveQR.Show
                   
End Select

End Sub



Private Sub Instellingen_Click(Index As Integer)

On Local Error Resume Next

Select Case Index
    Case 0
        'Teken.FLAGS = cdlPDPrintSetup
        Teken.PrinterDefault = True
        Teken.CancelError = False
        Teken.ShowPrinter
        Printer.EndDoc
        InstalleerAllePrinters
    Case 1
        VoorkeurDefinities.Show
    Case 2
        KeuzeBoxEditor.Show
    Case 3
        ProgrammaStruktuur.Show 1
        
    Case 4
        Msg = "Databasewerkwijze wijzigen.  Bent U zeker"
        KtrlBox = MsgBox(Msg, vbQuestion + vbDefaultButton2 + vbYesNo)
        If KtrlBox = vbYes Then
            Instellingen(4).Checked = Not Instellingen(4).Checked
            If Instellingen(4).Checked Then
                SnelHelp.Panels(3).text = "SQL"
            Else
                SnelHelp.Panels(3).text = "JET4"
            End If
            BeWaarTekst App.Title, "MSDE Gebruik", Instellingen(4).Checked
            AutoUnloadBedrijf
        End If
    
    Case 5
        ChDir App.path: MsgBox App.path & " wordt eerst als DEFAULT inhoudsopgave geplaatst", vbInformation
        Msg = "Kies JA en meest recent definitiebestand wordt automatisch uitgepakt." & vbCrLf & vbCrLf & "Kies 'NEE' om manueel de instructie arc xw marnt.arc *.* in DOS venster in te geven (VISTA en WINDOW 7 ADMINISTRATOR MODE)"
        KtrlBox = MsgBox(Msg, vbQuestion + vbDefaultButton3 + vbYesNoCancel)
        
        Select Case KtrlBox
            Case vbYes
                KtrlBox = Shell("arc xw marnt.arc *.*", 1)
            Case vbNo
                KtrlBox = Shell("COMMAND.COM", 3)
                MsgBox "Breng in DOS VENSTER manueel de instructie hieronder in" & vbCrLf & vbCrLf & " arc xw marnt.arc *.* " & vbCrLf & vbCrLf & "Na het uitpakken sluit het DOS venster af met instructie EXIT", vbInformation
            
            Case vbCancel
        End Select
            
End Select

End Sub



Private Sub MDIForm_Load()
Dim sVFL        As Integer
Dim T           As Integer
Dim A           As String
ReDim BCaption(5) As String
ReDim BColor(5) As Integer
Dim X           As Long
Dim FlTemp      As Integer
Dim Pipo        As Boolean
Dim BezigAanString As String

Dim path As String * 260
Dim ret As Integer
Dim Mailadres As String
Dim Paswoord As String

'Application Data lokatie ?
ret = SHGetFolderPath(0, CSIDL_APPDATA, 0, 0, path)

PROGRAM_LOCATION = App.path + "\"
Mim.Caption = appTitleAndVersion

usrMailAdres = "demo@rv.be"
usrPW = "9999"

ret = SHGetFolderPath(0, CSIDL_PERSONAL, 0, 0, path)
SYSTEM_MYPERSONALDOCUMENTS = Left(path, InStr(path, Chr(0)) - 1)

WaarBenIk:

BeWaarTekst App.Title, "VersionNumber", App.major & App.minor & Dec(App.Revision, "000")

LOCATION_MYDOCUMENTS = LaadTekst(App.Title, "Bedrijfsinhoudsopgave" & "2025")

BezigAanString = ""
BeWaarTekst "Programma", "LOCATION_", App.path

If Trim$(LOCATION_MYDOCUMENTS) <> "" Then
Else
    'Allereerste opstart of overgang
    'waar bevindt 'Mydocuments' zich ?
    ret = SHGetFolderPath(0, CSIDL_PERSONAL, 0, 0, path)
    GoSub MaakBewaarDataFolder
    If LOCATION_MYDOCUMENTS = "" Then
    Else
        GoTo WaarBenIk
    End If
End If

On Error GoTo MDIErrorOpvang
'Basis(3).Enabled = objUsers.CheckRights("LOGIN", 1)

Screen.MousePointer = vbHourglass
MIM_GLOBAL_DATE = Format(Now, "dd/mm/yyyy")
cmdWegBoekModus.AddItem "0: Geen BoekingsInfo tonen (EUROTEST niet actief)"
cmdWegBoekModus.AddItem "1: Enkel BoekingsInfo tonen bij EUR<>BEF verschil"
cmdWegBoekModus.AddItem "2: Altijd BoekingsInfo tonen"

BezigAanString = "VSOFT.INI instellingen inladen"
If Dir(PROGRAM_LOCATION + "vsoft.ini") = "" Then
    MsgBox "VSOFT.INI niet te vinden.  Installeer korrekt a.u.b."
    End
Else
    FlTemp = FreeFile
    Open PROGRAM_LOCATION + "vsoft.ini" For Input As FlTemp
    Do While Not EOF(FlTemp)
        Line Input #FlTemp, A
        Select Case LCase$(Left(A, 9))
            Case "programma"
            Case "assurnet "
                'tijdelijk ? bestand variabel gemaakt voor meerdere users..
                LOCATION_ASWEB = LTrim$(RTrim$(Mid(A, InStr(A, "=") + 1, Len(A) - InStr(A, "=")))) & Format(Now, "MMSS")
            Case "producent"
                ProducentNummer = LTrim$(RTrim$(Mid(A, InStr(A, "=") + 1, Len(A) - InStr(A, "="))))
        End Select
    Loop
    Close FlTemp
    LOCATION_ = LOCATION_MYDOCUMENTS + "\"
    PERIOD_FROMTO = ""
End If
BezigAanString = "MARNT.ARC proberen uitvergroten."
'If Dir(App.path + "\*.DEF") = "" Then
'    MsgBox PROGRAM_LOCATION + "marnt.arc Bestanden worden eerst uitvergroot onder DOS.  Druk straks OK pas wanneer de DOS taak VOLLEDIG BEEINDIGD werd !!"
'    X = Shell("arc xw marnt.arc *.*", 1)
'    MsgBox "Uitvergroten is beëindigd ?"
'End If
InitEerst

BezigAanString = "Vsoft licentieinfo venster openen."
Load Vsoft
'Vsoft.cmdLeesMij.Enabled = False
Vsoft.Ok.Default = True
Screen.MousePointer = vbNormal
Vsoft.Show 1
Vsoft.Refresh

On Error GoTo 0
BezigAanString = "Standaarddatabase openen."
Set adKBDB = New ADODB.Connection
Set adTBIB = New ADODB.Connection

adKBDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & PROGRAM_LOCATION & _
        "Default2022.mdb;"
adKBDB.Open

adTBIB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & PROGRAM_LOCATION & _
        "Telebib2.mdb;"
adTBIB.Open

Set adKBTable = New ADODB.Recordset
    adKBTable.CursorLocation = adUseServer
    adKBTable.Open "KeuzeBoxData", adKBDB, adOpenKeyset, adLockOptimistic, adCmdTableDirect '  adLockReadOnly, adCmdTableDirect
    adKBTable.Index = "BestandsNaam"

MIM_GLOBAL_DATE = Format(Now, "dd/mm/yyyy")

BezigAanString = "Virtuele fiches definieëren."
BCaption(1) = "Klanten"
BColor(1) = 9
BCaption(2) = "Leveranciers"
BColor(2) = 12
BCaption(3) = "Rekeningen"
BColor(3) = 15
BCaption(4) = "Artikels"
BColor(4) = 2

BezigAanString = "Invoerbox openen."
ntInputbox.Hide
BezigAanString = "Virtuele fiches maken."
For T = 1 To 3  '4 indien produktfiche
    BasisB(T).Caption = BCaption(T)
    BasisB(T).BackColor = QBColor(BColor(T))
    BasisB(T).Tag = Str$(T)
    'If T = 4 Then
    '    BasisB(T).WindowState = 0: BasisB(T).Height = BasisB(T).Height + 1600
    'End If
    BasisB(T).WindowState = 1
    BasisB(T).Enabled = False
Next

Screen.MousePointer = vbNormal
BezigAanString = "Kontrole Light versie."
If UCase(ProgrammaVersie) = UCase("NT Light versie") Then
    Mim.Basis(1).Enabled = False
End If
BezigAanString = "Standaardinstellingen inladen"
XDoEvents = DoEvents
BezigAanString = "Hoofdscherm tonen"
If LaadTekst(App.Title, "TaakBalkZichtbaar") = "" Then
    BeWaarTekst App.Title, "TaakBalkZichtbaar", True
End If
Mim.tbToolBar.Visible = LaadTekst(App.Title, "TaakBalkZichtbaar")
If Mim.tbToolBar.Visible = False Then Mim.Basis(8).Checked = False
Mim.Show
LaadInstellingen Me
Mim.Arrange 3
BezigAanString = "Programmalicentie info verbergen"
Vsoft.Hide

BezigAanString = "Printers installeren."
If InStr(UCase(Command), "/P=FALSE") Then
Else
    InstalleerAllePrinters
End If

Dim ZoekHTM As Variant
Dim ZoekNietMeerHTM As Variant

BedrijfOpenen.Show
ZoekHTM = App.path & "\LeesMij" & App.major & App.minor & Format(App.Revision, "000") & ".htm"
ZoekNietMeerHTM = App.path & "\LeesMij" & App.major & App.minor & App.Revision & ".not"

Dim r As Long

If Dir(ZoekHTM) <> "" Then
    If Dir(ZoekNietMeerHTM) <> "" Then
    Else
        KtrlBox = MsgBox("Bij volgende opstart LeesMij.htm voor deze versie en revisie verbergen", vbQuestion + vbYesNo + vbDefaultButton2, App.Title & " Versie " & App.major & "." & App.minor & "." & App.Revision)
        If KtrlBox = vbYes Then
            ScrMaakTekstBestand (Now), (ZoekNietMeerHTM)
        End If
        Err = 0
        On Error Resume Next
        If ShellExecuteWithFallback(ZoekHTM) = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
    End If
End If
Mim.DatumVerwerking.Value = Now
Exit Sub

MDIErrorOpvang:
      Msg = "Stopcode : " + Str$(Err) + vbCrLf + vbCrLf
Msg = Msg + "Bezig aan :" + BezigAanString + vbCrLf + vbCrLf
Msg = Msg + "Visual Basic 6.x Enterprise Edition Error:" + vbCrLf
Msg = Msg + Error + vbCrLf + vbCrLf
Msg = Msg + "Voor meer inlichtingen neem contact op met R&VSOFT 053/21.59.25 !"
MsgBox Msg, 0, "Onverwachte situatie"
End

MaakBewaarDataFolder:
LOCATION_MYDOCUMENTS = Left(path, InStr(path, Chr(0)) - 1)
If InStr(UCase(LOCATION_MYDOCUMENTS), "ONEDRIVE") Then
    Msg = "Vanaf Windows 11 wordt voorgesteld uw documenten op OneDrive te bewaren." & vbCrLf & vbCrLf
    Msg = Msg & "Voor uw boekhouddata kan dat indien U de enige gebruiker bent." & vbCrLf
    Msg = Msg & "en niemand anders via netwerkverbinding de data dient te bewerken" & vbCrLf & vbCrLf
    Msg = Msg & "Meerdere gebruikers opteren beter om manueel marnt\data op een gedeelde locatie te plaatsen"
    MsgBox Msg, vbExclamation, "OneDrive gevonden!"
End If

Msg = "Nieuwe folder " & LOCATION_MYDOCUMENTS & "\marNT\Data" & vbCr
Msg = Msg & "moet pér 15/08/2025 KUNNEN aangemaakt worden" & vbCr & vbCr
Msg = Msg & "hierna aanmaken?" & vbCr & vbCr & "Kies nee om manueel in te stellen"
KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
If KtrlBox = vbYes Then
    If scrFolderBestaat(LOCATION_MYDOCUMENTS + "\marNT\Data") Then
    Else
        On Error Resume Next
        Err = 0
        fs.CreateFolder (LOCATION_MYDOCUMENTS + "\marNT")
        If Err Then
            Msg = "Folder '" & LOCATION_MYDOCUMENTS & "\marNT' voor dit toestel kan niet gecontroleerd/gecreëerd worden (mogelijk schrijfbeveiligd en/of bevoegdheidsbeperkingen binnen netwerkomgeving)." & vbCr & vbCr & "Gelieve de Microsoft aanbevelingen rondom gebruik van folder 'Mijn dokumenten' te respecteren." & vbCr & vbCr & "Raadpleeg de beheerder van uw syste(e)m(en) voor rechzettingen indien nodig."
            MsgBox Error & vbCr & vbCr & Msg, vbExclamation
            Return
        End If
        fs.CreateFolder (LOCATION_MYDOCUMENTS + "\marNT\Data")
        If Err Then
            Msg = "Folder 'Mijn documenten' voor uw toestel is schrijfbeveiligd (vermoedelijk binnen netwerkomgeving).  U dient zelf de nodige rechtzettingen uit te voeren.  Raadpleeg de beheerder van uw syste(e)m(en) voor de nodige rechzettingen indien nodig."
            MsgBox Error & vbCr & vbCr & Msg, vbExclamation
            Return
        End If
    End If
    LOCATION_MYDOCUMENTS = LOCATION_MYDOCUMENTS + "\marnt\data"
Else
    LOCATION_MYDOCUMENTS = ""
End If
BeWaarTekst App.Title, "Bedrijfsinhoudsopgave" & "2025", LOCATION_MYDOCUMENTS
Return
            
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)

Dim X As Boolean
Dim manualFiles As String
Dim strAddress As String
Dim r As Long

manualFiles = ""

TotalClose
If WindowState = vbMinimized Then
    WindowState = vbNormal
End If
X = InstellingenBewaard(Me)

If LaadTekst("dnnInstellingen", "Mario") = "" Then
Else
    strAddress = LaadTekst("dnnInstellingen", "Mario")
    manualFiles = Dir(strAddress & "\*.*")
    If manualFiles = "" Then
    Else
        If ShellExecuteWithFallback(strAddress) = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
        DoEvents
    End If
End If
End

End Sub


Private Sub mnuXML_Click(Index As Integer)

Select Case Index
    Case 0
        'frmADOXML.Show
    Case 1
        'frmXMLbackup.Show
        'Mim.Arrange 0
        
End Select

End Sub

Private Sub msWindows_Click(Index As Integer)

'Dim frmB As New frmBrowser
On Local Error Resume Next
Err = 0
Select Case Index
    Case 0
        If Err Then MsgBox Error: Exit Sub
        'frmB.StartingAddress = "http://windowsupdate.microsoft.com"
        'If Err Then MsgBox Error: Exit Sub
        'frmB.Show
        'frmB.WindowState = vbMaximized
        'If Err Then MsgBox Error: Exit Sub
    
    Case 1
        If Err Then MsgBox Error: Exit Sub
        'frmB.StartingAddress = "http://support.microsoft.com/servicedesks/fileversion/dllinfo.asp?fr=0&sd=tech"
        'If Err Then MsgBox Error: Exit Sub
        'frmB.Show
        'frmB.WindowState = vbMaximized
        'If Err Then MsgBox Error: Exit Sub
        
    Case 2
        If Err Then MsgBox Error: Exit Sub
        'frmB.StartingAddress = "http://office.microsoft.com/default.aspx"
        'If Err Then MsgBox Error: Exit Sub
        'frmB.Show
        'frmB.WindowState = vbMaximized
        'If Err Then MsgBox Error: Exit Sub
        
End Select

End Sub

Private Sub Overgang_Click(Index As Integer)
Dim XX As Integer

Select Case Index
    Case 0
        NieuwBoekjaar
        Basis_Click 0
    Case 1
        If BJPERDAT!Boekjaar.ListIndex = 0 Then
            MsgBox "Journaalbestand kan nooit opgekuist worden met het hoogste boekjaar actief.  Aktiveer een lager boekjaar a.u.b. !"
            BJPERDAT.WindowState = 0
            BJPERDAT!Boekjaar.SetFocus
            Exit Sub
        End If
        Msg = ""
        If BJPERDAT!Boekjaar.ListIndex = 1 Then
            Msg = Msg + "OPGELET :" + vbCrLf + "Het tweede hoogste boekjaar is actief en "
            Msg = Msg + "er zijn " + Format(BJPERDAT!Boekjaar.ListCount) + " boekjaren." + vbCrLf + vbCrLf
            Msg = Msg + "ADVIES :" + vbCrLf + "Normaal gezien laat U best ook het tweedelaatste boekjaar bestaan..." + vbCrLf
            Msg = Msg + "Voor statistisch onderzoek (grafisch/cijfermatig) best niets vernietigen!!" + vbCrLf + vbCrLf

        End If
        Msg = Msg + "ALLE journaallijnen met datum tot en met "
        Msg = Msg + Right(BJPERDAT!PeriodeBoekjaar.List(BJPERDAT!PeriodeBoekjaar.ListCount - 1), 10)
        Msg = Msg + " vernietigen !" + vbCrLf + vbCrLf
        Msg = Msg + "Bent U zeker ?"
        Ktrl = MsgBox(Msg, vbYesNo + vbDefaultButton2, "Opkuis journaalbestanden")
        If Ktrl = vbYes Then
            BJPERDAT!PeriodeBoekjaar.ListIndex = BJPERDAT!PeriodeBoekjaar.ListCount - 1
            bFirst TABLE_JOURNAL, 0
            If Ktrl Then
            Else
                Screen.MousePointer = vbHourglass
                Mim.Enabled = False
                bBegin
                Do
                    If Right(KEY_BUF(TABLE_JOURNAL), 8) <= Right(PERIOD_FROMTO, 8) Then
                        SnelHelpPrint KEY_BUF(TABLE_JOURNAL), BL_LOGGING
                        bDelete TABLE_JOURNAL
                    End If
                    bNext TABLE_JOURNAL
                    XDoEvents = DoEvents
                    If Ktrl Then
                        Exit Do
                    End If
                Loop
                bEnd
                Screen.MousePointer = vbNormal
                Mim.Enabled = True
                Msg = "ALLE aankoop- en verkoopdokumenten met datum tot en met "
                Msg = Msg + Right(BJPERDAT!PeriodeBoekjaar.List(BJPERDAT!PeriodeBoekjaar.ListCount - 1), 10)
                Msg = Msg + " kan U hierna eveneens verwijderen door één dergelijk dokument via klanten- of leveranciersfiche ter vernietiging aan te duiden"
                MsgBox Msg
            End If
        End If
        
End Select

End Sub

Private Sub pBeheer_Click(Index As Integer)

On Local Error Resume Next

Select Case Index
    Case 0
        Venster.Show
    Case 1
        KwijtingDrukken.Show
    Case 2
        GridText = ""
        KwijtingBoeken.Show
    Case 4
        AS1VrijBericht.Show
    
    Case 6
        'KTRLMijStatistiek "912"
    Case 7
        MijStat.Show 1
    Case 8
        As2TelebibIn.Show
        As2TelebibIn.SetFocus
    Case 10
        frmGroeneKaart.Show
    Case 11
        Dim r As Long
        If ShellExecuteWithFallback("https://cdd.vsoft.be") = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
        
End Select

End Sub



Private Sub Peppol_Click(Index As Integer)

'Dim frmB As New frmBrowser
Dim dKtrl As Double
Dim path As String * 260
Dim deFolder As String
Dim deSubFolder As String

Dim ret As Integer
Dim r As Long

On Local Error Resume Next
Err = 0
Select Case Index
    Case 0
        If ShellExecuteWithFallback("https://peppol-tools.ademico-software.com/ui/document-validator") = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If

    Case 1
        If ShellExecuteWithFallback("https://docs.peppol.eu/poacc/billing/3.0/") = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If

    Case 2
        ' Voorbeeld gebruik
        If ShellExecuteWithFallback("https://clickonce.vsoft.be/MarSync/publish.htm") = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
End Select
    
End Sub

Private Sub SchikMenu_Click(Index As Integer)

Select Case Index
    Case 0
        Mim.Arrange 0
    Case 1
        Mim.Arrange 1
    Case 2
        Mim.Arrange 2
    Case 3
        Mim.Arrange 3
End Select
XDoEvents = DoEvents

End Sub

Private Sub TotalClose()

AutoUnloadBedrijf
For COUNT_TO = 0 To 9
    Set rsMAR(COUNT_TO) = Nothing
Next
adKBDB.Close
Set adKBDB = Nothing

On Error Resume Next
adntDB.Close
If Err Then Exit Sub
ntDB.Close

Set adntDB = Nothing
Set ntDB = Nothing


End Sub





Private Sub SnelHelp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dPip As Double
Dim sPip As String * 2
    
Mim.Instellingen(4).Enabled = True

If Button = 2 Then
    Msg = InputBox("Vervolledig de toegangskode in overleg met" + vbCrLf + "Roelandt & Vermoesen (tel. 053/21.59.25)", "Extra kontrole marInteraal !", Format(Now, "YYYYMMDD"))
    If Msg = "" Then
        Exit Sub
    ElseIf Msg = "435605026" Then
        GoTo Foefke
    ElseIf Msg = "thequickbrownfox" Then
        VSF_PRO = True
        Exit Sub
    End If
    dPip = Val(Left(Msg, 10))
    sPip = Format(dPip - Int(dPip / 97) * 97, "00")
    If sPip <> Right(Msg, 2) Then
        MsgBox "Ongeldige toegangskode."
    ElseIf Mid(Msg, 9, 2) <> "90" Then
        MsgBox "Ongeldige toegangskode."
        Exit Sub
    Else
        VSF_PRO = True
    End If
    Exit Sub

Foefke:
    Select Case Msg
        Case "435605026"
            X = KtrlTVSOMERS
        Case Else
            MsgBox "Neem telefonisch kontakt op met ROELANDT & VERMOESEN voor meer inlichtingen rondom de extra kontroleroutine's.  Bel 053/21.59.25"
    End Select
End If

End Sub

Private Sub SnelHelpTijd_Timer()

SnelHelpTijd.Enabled = False
Mim.SnelHelp.Panels(4).text = ""

If LOG_PRINT = "" Then
Else
    MsgBox LOG_PRINT
    LOG_PRINT = ""
End If

End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As Button)
    
'Dim frmB As frmBrowser
Dim r As Long
Dim strAdress As String

Err = 0
On Error Resume Next
     
    Select Case Button.Key
        Case "Open"
           Basis_Click 0
        
        Case "SQLZoek"
            COUNT_TO = adntDB.Version
            If Err Then
                Beep
            Else
                frmADOSQLOpzoeker.Show
                'frmOntwerpSQLVenster.Show
            End If
                
        Case "computerbeheer"
        
            'On Error GoTo 0
            'KtrlBox = Shell("%SystemRoot%\system32\compmgmt.msc /s", vbNormalFocus)
            
        Case "vsoft"
            
            If LaadTekst("dnnInstellingen", "Mario") = "" Then
                MsgBox "Eerst uw URL instellen a.u.b. probeer later opnieuw", vbExclamation
                frmDNNinstellingen.Show 1
                Exit Sub
            End If
            
            strAdress = LaadTekst("dnnInstellingen", "Mario")
            If ShellExecuteWithFallback(strAdress) = False Then
                MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
            End If
                   
        Case "server1"
            
            'Set frmB = New frmBrowser
            
            If Err Then MsgBox Error: Exit Sub
            'frmB.StartingAddress = "http://rvsoft.no-ip.biz/rvDNN"
            
            'frmB.StartingAddress = "http://www.rv.be/Services/mar/Default.aspx?rvPW=" & usrPW & "&rvA100=" & usrMailAdres
            'frmB.WindowState = vbMaximized
            
            If Err Then MsgBox Error: Exit Sub
            'frmB.Show
            If Err Then MsgBox Error: Exit Sub
                   
        Case "server2"
            
            If Err Then MsgBox Error: Exit Sub
            If LaadTekst("dnnInstellingen", "Archief") = "" Then
                MsgBox "Eerst uw URL instellen a.u.b. probeer later opnieuw", vbExclamation
                frmDNNinstellingen.Show 1
            Exit Sub
        End If
                
        strAdress = LaadTekst("dnnInstellingen", "Archief")
        If ShellExecuteWithFallback(strAdress) = False Then
                MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
            End If
                   
        Case "sepaconverter"
            'SEPA.Show
                   
        Case Else
            MsgBox Button.Key + " niet toegankelijk"
        End Select
    
End Sub

Sub adoSQLOmwerking()
    
    Dim rtKode As Boolean

    'Aanmaak van de Tabellen
    'rtKode = adxMaakTabel("Verkoopdokumenten")
    '   rtKode = adxMaakTabel("DetailVerkoop")
    
    'rtKode = adxMaakTabel("Bestelbons")
    '    rtKode = adxMaakTabel("DetailBestelBons")
        
    'rtKode = adxMaakTabel("Bestelbons")
    '    rtKode = adxMaakTabel("DetailOffertes")

    'Aanmaak van de Kolommen

End Sub

Private Function KontroleerDatabase()

Dim Aantal As Long
Dim Veldje As String

Dim eVeldje As String
Dim bVeldje As String

Dim rcDummy As ADODB.Recordset

Dim Speciaal As Boolean
Dim AVktrl As Boolean

'Journaallijnen met Null in bedrag verwijderen ?
Msg = "Overbodige journaallijnen met 'Null' waarde in bedrag verwijderen."
KtrlBox = MsgBox(Msg, vbYesNo + vbDefaultButton2 + vbQuestion)
If KtrlBox = vbYes Then
    On Error Resume Next
    Err = 0
    Screen.MousePointer = vbHourglass
    adntDB.Execute "DELETE * FROM Journalen WHERE isNull(v068)", Aantal
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox Error
    ElseIf Aantal Then
        MsgBox "Er zijn " & Aantal & " overbodige journaallijnen verwijderd", vbInformation
    Else
        MsgBox "Er was niets te verbeteren.", vbInformation
    End If
End If

'Diverse Journaalposten jaartal 'jj' naar 'jjjj'
Msg = "Diverse posten met jaartal 'jj' naar 'jjjj' verbeteren"
KtrlBox = MsgBox(Msg, vbYesNo + vbDefaultButton2 + vbQuestion)
If KtrlBox = vbYes Then
    On Error Resume Next
    Err = 0
    Screen.MousePointer = vbHourglass
    adntDB.Execute "UPDATE Journalen SET v033 = Mid(v033,1,2) + v066 WHERE Left(v033,1)=" + Chr(34) + "D" + Chr(34) + " AND Len(Trim$(v033)) = 8", Aantal
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox Error
    ElseIf Aantal Then
        MsgBox "Er zijn " & Aantal & " journaallijnen verbeterd", vbInformation
    Else
        MsgBox "Er was niets te verbeteren.", vbInformation
    End If
End If

'BEF Jaarsaldo's naar EUR jaarsaldo's omvormen
Msg = "Alle Jaarsaldo's Algemene Rekeningen in EUR beschikbaar maken als DECIMAAL (Microsoft.NET) veld"
KtrlBox = MsgBox(Msg, vbYesNo + vbDefaultButton1 + vbQuestion)
If KtrlBox = vbYes Then
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    Set rcDummy = New ADODB.Recordset
    Err = 0
    Msg = "SELECT * FROM Rekeningen"
    rcDummy.CursorLocation = adUseClient
    rcDummy.Open Msg, adntDB, adOpenForwardOnly, adLockOptimistic
    If Err Then MsgBox "stop"
    rcDummy.MoveFirst
    On Error GoTo 0
    Do While Not rcDummy.EOF
        For COUNT_TO = 22 To 31
            bVeldje = "e0" + Dec((COUNT_TO), "00")
            eVeldje = "dece0" + Dec((COUNT_TO), "00")
            Err = 0
            On Error Resume Next
            rcDummy(eVeldje) = Val(rcDummy(bVeldje))
            If Err Then rcDummy(eVeldje) = 0
            SnelHelpPrint rcDummy("v019") & rcDummy(eVeldje), BL_LOGGING
        Next
        rcDummy.MoveNext
    Loop
    Screen.MousePointer = vbNormal
    MsgBox "Einde update DECIMAAL veld(en) rekeningen"
    rcDummy.Close
    Set rcDummy = Nothing
End If

End Function

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Mim.Report.IsOpen = True Then
        MsgBox "Sluit eerst het Verkoopvenster a.u.b.", vbInformation
        Cancel = True
    End If
    
    Dim sPath As String
    Dim sFile As String
    
    If LOCATION_COMPANYDATA = "" Then
    Else
        sPath = LOCATION_COMPANYDATA & "peppol\in\"
        If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
        sFile = Dir$(sPath & "*.pdf")          ' first match
        Do While sFile <> ""
            Err = 0
            On Local Error Resume Next
            Kill sPath & sFile
        sFile = Dir$() ' subsequent matches
        Loop

        sPath = LOCATION_COMPANYDATA & "peppol\out\"
        If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    
        sFile = Dir$(sPath & "*.pdf")          ' first match
        Do While sFile <> ""
            Err = 0
            On Local Error Resume Next
            Kill sPath & sFile
        sFile = Dir$() ' subsequent matches
        Loop
    End If
    
End Sub

Private Sub DetectClickOnceShortcut()
    Dim startMenuPath As String
    Dim shortcutFile As String

    DoEvents
    ' Path to the ClickOnce Start Menu folder
    startMenuPath = Environ("APPDATA") & "\Microsoft\Windows\Start Menu\Programs\Vsoft Administratieve Software"

    ' Search for the .appref-ms file
    shortcutFile = Dir(startMenuPath & "\*.appref-ms")

    If shortcutFile <> "" Then
        'MsgBox "ClickOnce shortcut found: " & startMenuPath & "\" & shortcutFile, vbInformation, "Shortcut Detected"
        If ShellExecuteWithFallback(startMenuPath & "\" & shortcutFile) = False Then
            MsgBox "Kon bestand niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
        End If
    Else
        MsgBox "ClickOnce snelkoppeling voor MarSync niet gevonden in:" & vbCrLf & vbCrLf & startMenuPath, vbExclamation, "Snelkoppeling Niet Gevonden"
    End If
    
End Sub
