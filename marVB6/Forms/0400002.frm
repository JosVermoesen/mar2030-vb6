VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form InventarisOpgave 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventaris Stock"
   ClientHeight    =   3870
   ClientLeft      =   3630
   ClientTop       =   1530
   ClientWidth     =   4320
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3870
   ScaleWidth      =   4320
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TekstLijn 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   3120
      TabIndex        =   14
      Top             =   1080
      Width           =   1080
   End
   Begin VB.CheckBox cbNegeerRest 
      Caption         =   "Negeer restwaarde < EUR"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton cmdKommaCorrectie 
      Caption         =   "Decimale Komma verbetering"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Alignment       =   1  'Right Justify
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   2400
      TabIndex        =   10
      Top             =   3570
      Width           =   1830
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Eindinventaris ZONDER kontrole !"
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
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3060
      Width           =   2040
   End
   Begin VB.CommandButton Fase 
      Caption         =   "Fase"
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
      Left            =   3000
      TabIndex        =   8
      Top             =   1440
      Width           =   1200
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
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   690
      Width           =   2415
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   0
      Top             =   90
      Width           =   3000
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   2640
      TabIndex        =   1
      Top             =   390
      Width           =   1560
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Afdrukken"
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
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   1200
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
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid X 
      Height          =   1605
      Left            =   180
      TabIndex        =   11
      Top             =   1410
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   2831
      _Version        =   393216
      Rows            =   6
      BackColor       =   -2147483624
      SelectionMode   =   1
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
   Begin VB.Label Label1 
      Caption         =   "Sortering"
      Height          =   240
      Index           =   2
      Left            =   210
      TabIndex        =   7
      Top             =   720
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "SubTitel"
      Height          =   240
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Datum"
      Height          =   240
      Index           =   1
      Left            =   210
      TabIndex        =   5
      Top             =   420
      Width           =   1005
   End
End
Attribute VB_Name = "InventarisOpgave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim psTekst(5) As String
Dim VeldTXT(14) As String

Dim aa                As String

Dim ProduktOK             As Long
Dim ProduktNOK            As Long
Dim ProduktZOK            As Long
Dim WaardeNegeer        As Long

Dim TLijnen             As Long

Dim TotaalBeginvoorraad     As Currency
Dim TotaalEindvoorraad     As Currency

Dim BedragAankoop As Currency
Dim BedragVerkoop  As Currency
Dim BedragBeginVoorraad As Currency

Dim AlgTotBegin     As Double
Dim AlgTotAankoop   As Double
Dim AlgTotVerkoop   As Double
Dim AlgTotEindVoor  As Double

Private Sub Annuleren_Click()

Unload Me

End Sub



Private Sub cmdKommaCorrectie_Click()

Dim Aantal As Long

'Veld J (Minimumstock)
Err = 0
On Error Resume Next
Msg = "UPDATE Produkten SET v115 = Left(Str$(Cdbl(v115)),10) WHERE InStr$(v115," + Chr(34) + "," + Chr(34) + ") <> 0"
Screen.MousePointer = vbHourglass
adntDB.Execute Msg, Aantal

Screen.MousePointer = vbNormal
If Err Then
    MsgBox Error
Else
    Msg = Msg + vbCr + vbCr
    Msg = Msg + Str$(Aantal) + " zijn er bijgewerkt."
    MsgBox Msg, , "Met succes: Aantal Minimumstock"
End If

'Veld N (Aantal Aankoop)
Err = 0
On Error Resume Next
Msg = "UPDATE Produkten SET v119 = Left(Str$(Cdbl(v119)),10) WHERE InStr$(v119," + Chr(34) + "," + Chr(34) + ") <> 0"
Screen.MousePointer = vbHourglass
adntDB.Execute Msg, Aantal
Screen.MousePointer = vbNormal
If Err Then
    MsgBox Error
Else
    Msg = Msg + vbCr + vbCr
    Msg = Msg + Str$(Aantal) + " zijn er bijgewerkt."
    MsgBox Msg, , "Met succes: Aantal Aangekocht"
End If

'Veld 0 (Aantal Verkoop)
Err = 0
On Error Resume Next
Msg = "UPDATE Produkten SET v120 = Left(Str$(Cdbl(v120)),10) WHERE InStr$(v120," + Chr(34) + "," + Chr(34) + ") <> 0"
Screen.MousePointer = vbHourglass
adntDB.Execute Msg, Aantal
Screen.MousePointer = vbNormal
If Err Then
    MsgBox Error
Else
    Msg = Msg + vbCr + vbCr
    Msg = Msg + Str$(Aantal) + " zijn er bijgewerkt."
    MsgBox Msg, , "Met succes: Aantal Verkocht"
End If

'Veld P (Aantal Aankoop)
Err = 0
On Error Resume Next
Msg = "UPDATE Produkten SET v114 = Left(Str$(Cdbl(v114)),10) WHERE InStr$(v114," + Chr(34) + "," + Chr(34) + ") <> 0"
Screen.MousePointer = vbHourglass
adntDB.Execute Msg, Aantal
Screen.MousePointer = vbNormal
If Err Then
    MsgBox Error
Else
    Msg = Msg + vbCr + vbCr
    Msg = Msg + Str$(Aantal) + " zijn er bijgewerkt."
    MsgBox Msg, , "Met succes: Aantal Beginvoorraad"
End If

'Veld H (Bedrag Verkoop)
Err = 0
On Error Resume Next
Msg = "UPDATE Produkten SET v112 = Left(Str$(Cdbl(v112)),15) WHERE InStr$(v112," + Chr(34) + "," + Chr(34) + ") <> 0"
Screen.MousePointer = vbHourglass
adntDB.Execute Msg, Aantal
Screen.MousePointer = vbNormal
If Err Then
    MsgBox Error
Else
    Msg = Msg + vbCr + vbCr
    Msg = Msg + Str$(Aantal) + " zijn er bijgewerkt."
    MsgBox Msg, , "Met succes: Bedrag Verkoop"
End If
        
'Veld I (Bedrag Aankoop)
Err = 0
On Error Resume Next
Msg = "UPDATE Produkten SET v113 = Left(Str$(Cdbl(v113)),15) WHERE InStr$(v113," + Chr(34) + "," + Chr(34) + ") <> 0"
Screen.MousePointer = vbHourglass
adntDB.Execute Msg, Aantal
Screen.MousePointer = vbNormal
If Err Then
    MsgBox Error
Else
    Msg = Msg + vbCr + vbCr
    Msg = Msg + Str$(Aantal) + " zijn er bijgewerkt."
    MsgBox Msg, , "Met succes: Bedrag aankoop"
End If

'Veld Q (Bedrag jaaraankoop)
Err = 0
On Error Resume Next
Msg = "UPDATE Produkten SET v121 = Left(Str$(Cdbl(v121)),10) WHERE InStr$(v121," + Chr(34) + "," + Chr(34) + ") <> 0"
Screen.MousePointer = vbHourglass
adntDB.Execute Msg, Aantal
Screen.MousePointer = vbNormal
If Err Then
    MsgBox Error
Else
    Msg = Msg + vbCr + vbCr
    Msg = Msg + Str$(Aantal) + " zijn er bijgewerkt."
    MsgBox Msg, , "Met succes: Bedrag Jaaraankoop"
End If

'Veld R (Bedrag jaarverkoop)
Err = 0
On Error Resume Next
Msg = "UPDATE Produkten SET v122 = Left(Str$(Cdbl(v122)),10) WHERE InStr$(v122," + Chr(34) + "," + Chr(34) + ") <> 0"
Screen.MousePointer = vbHourglass
adntDB.Execute Msg, Aantal
Screen.MousePointer = vbNormal
If Err Then
    MsgBox Error
Else
    Msg = Msg + vbCr + vbCr
    Msg = Msg + Str$(Aantal) + " zijn er bijgewerkt."
    MsgBox Msg, , "Met succes: Bedrag Jaarverkoop"
End If

'Veld S (Bedrag beginvoorraad)
Err = 0
On Error Resume Next
Msg = "UPDATE Produkten SET v123 = Left(Str$(Cdbl(v123)),10) WHERE InStr$(v123," + Chr(34) + "," + Chr(34) + ") <> 0"
Screen.MousePointer = vbHourglass
adntDB.Execute Msg, Aantal
Screen.MousePointer = vbNormal
If Err Then
    MsgBox Error
Else
    Msg = Msg + vbCr + vbCr
    Msg = Msg + Str$(Aantal) + " zijn er bijgewerkt."
    MsgBox Msg, , "Met succes: Bedrag beginvoorraad"
End If

End Sub

Private Sub Command1_Click()

LijstDrukken 1

End Sub


Private Sub Drukken_Click()

LijstDrukken 1

End Sub

Private Sub Fase_Click()
Dim CumulBedrag As Double
Dim Instock As Double

ProduktOK = 0
ProduktNOK = 0
ProduktZOK = 0
WaardeNegeer = 0

bFirst TABLE_PRODUCTS, Val(Left(Sortering.Text, 2))
If Ktrl Then
    Beep
    Exit Sub
Else
    Screen.MousePointer = vbHourglass
    GoSub DoeDeKontrole
End If

Do
    bNext TABLE_PRODUCTS
    If Ktrl Then
        Exit Do
    Else
        GoSub DoeDeKontrole
    End If
Loop
Screen.MousePointer = vbNormal
If ProduktNOK Then
    Msg = "Er zijn " + Format(ProduktNOK) + " artikels niet in orde.  Kontrolelijst afdrukken ?"
    Ktrl = MsgBox(Msg, 292)
    If Ktrl = 6 Then
        LijstDrukken 0
    End If
ElseIf ProduktOK <> 0 Then
    Drukken.Enabled = True
    Fase.Visible = False
End If
Exit Sub

DoeDeKontrole:
RecordToVeld TABLE_PRODUCTS
If vBibTekst(TABLE_PRODUCTS, "#v108 #") = "D" Or vBibTekst(TABLE_PRODUCTS, "#v108 #") = "S" Then
    Return
Else
    VeldTXT(0) = vSet(vBibTekst(TABLE_PRODUCTS, "#v102 #"), 13)  'artikelnummer
    VeldTXT(1) = vSet(vBibTekst(TABLE_PRODUCTS, "#v105 #"), 30)  'naam
    SnelHelpPrint VeldTXT(0), BL_LOGGING
    
    VeldTXT(2) = vSet(vBibTekst(TABLE_PRODUCTS, "#v118 #"), 7) 'rekeningvoorraad
    
    VeldTXT(5) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v119 #")), MASK_SY(7))    'eenheden AK
    VeldTXT(7) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v120 #")), MASK_SY(7))    'eenheden VK
    VeldTXT(3) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v114 #")), MASK_SY(7))  'eenheden BV
    
    If bhEuro Then
        VeldTXT(6) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e121 #")), Right(MASK_SY(4), 8))   'waardeaankoop
        VeldTXT(8) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e122 #")), Right(MASK_SY(4), 8))    'waardeverkoop
        VeldTXT(4) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e123 #")), Right(MASK_SY(4), 8))    'waardebeginvoorraad
        VeldTXT(9) = Dec$(Val(VeldTXT(3)) + Val(VeldTXT(5)) - Val(VeldTXT(7)), MASK_SY(7))  'eindstock
        VeldTXT(10) = Dec$(Val(VeldTXT(4)) + Val(VeldTXT(6)) - Val(VeldTXT(8)), Right(MASK_SY(4), 8))  'eindwaarde
    Else
        VeldTXT(6) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v121 #")), MASK_SY(8))    'waardeaankoop
        VeldTXT(8) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v122 #")), MASK_SY(8))    'waardeverkoop
        VeldTXT(4) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v123 #")), MASK_SY(8))    'waardebeginvoorraad
        VeldTXT(9) = Dec$(Val(VeldTXT(3)) + Val(VeldTXT(5)) - Val(VeldTXT(7)), MASK_SY(7))  'eindstock
        VeldTXT(10) = Dec$(Val(VeldTXT(4)) + Val(VeldTXT(6)) - Val(VeldTXT(8)), MASK_SY(8))  'eindwaarde
    End If
    
    If Val(VeldTXT(10)) = 0 And Val(VeldTXT(9)) <> 0 Then  'klopt iets niet
        ProduktNOK = ProduktNOK + 1
        GoSub GridBijwerken
    ElseIf Val(VeldTXT(9)) = 0 And Val(VeldTXT(10)) <> 0 Then 'idem
        If cbNegeerRest.Value = 1 Then
            If bhEuro Then
                If Abs(Val(VeldTXT(10))) < Val(TekstLijn(2).Text) Then
                    WaardeNegeer = WaardeNegeer + 1
                Else
                    ProduktNOK = ProduktNOK + 1
                End If
            Else
                If Abs(Val(VeldTXT(10)) / EURO) < Val(TekstLijn(2).Text) Then
                    WaardeNegeer = WaardeNegeer + 1
                Else
                    ProduktNOK = ProduktNOK + 1
                End If
            End If
        Else
            ProduktNOK = ProduktNOK + 1
        End If
        GoSub GridBijwerken
    ElseIf Val(VeldTXT(9)) < 0 Or Val(VeldTXT(10)) < 0 Then
        ProduktNOK = ProduktNOK + 1
        GoSub GridBijwerken
    Else
        ProduktOK = ProduktOK + 1
        GoSub GridBijwerken
    End If
End If
Return

GridBijwerken:
X.Col = 1
X.Row = 1
X.Text = Format(ProduktNOK)
X.Row = 2
X.Text = Format(ProduktOK)
X.Row = 3
For COUNT_TO = 3 To 8
    If Val(VeldTXT(COUNT_TO)) <> 0 Then
        ProduktZOK = ProduktZOK + 1
        X.Text = Format(ProduktZOK)
        Exit For
    End If
Next
X.Row = 4
X.Text = Format(WaardeNegeer)
X.Refresh

Return

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Dim T As Integer
Dim TempB_MODUS As Integer

TekstLijn(0).Text = ""
TekstLijn(1).Text = MIM_GLOBAL_DATE

X.Row = 0
X.Col = 0
X.Text = "Status"
X.ColWidth(0) = 1400
X.ColAlignment(0) = flexAlignGeneral

X.Col = 1
X.Text = "Waarde"
X.ColWidth(1) = 1000
X.ColAlignment(1) = flexAlignGeneral

X.Row = 1
X.Col = 0
X.Text = "Te kontroleren"

X.Row = 2
X.Text = "In orde"

X.Row = 3
X.Text = "Effectief nuttig"

X.Row = 4
X.Text = "Overslag Rest"

TekstLijn(2).Text = "0.03"
Sortering.Clear
For T = 0 To FL_NUMBEROFINDEXEN(TABLE_PRODUCTS)
    Sortering.AddItem Format(T, "00") + ":" + FLINDEX_CAPTION(TABLE_PRODUCTS, T)
Next
Sortering.ListIndex = 0
Fase.Caption = "Fase 1"

If String99(READING, 62) = "1" Then
    MsgBox "Eindinventaris reeds afgesloten !"
    Fase.Enabled = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

bClose TABLE_LEDGERACCOUNTS
bClose TABLE_PRODUCTS

End Sub

Private Sub InitVelden()
Dim T As Integer

REPORT_FIELD(0) = "Kodenummer"
    REPORT_TAB(0) = 1

REPORT_FIELD(1) = "Naam/Omschrijving 30 tekens"
    REPORT_TAB(1) = 15

REPORT_FIELD(2) = "BH.Rek."
    REPORT_TAB(2) = 46

REPORT_FIELD(3) = "BV Eenh."
    REPORT_TAB(3) = 55

REPORT_FIELD(4) = "  Waarde"
    REPORT_TAB(4) = 64

REPORT_FIELD(5) = "AK.Eenh."
    REPORT_TAB(5) = 74

REPORT_FIELD(6) = "  Waarde"
    REPORT_TAB(6) = 83

REPORT_FIELD(7) = "VK.Eenh."
    REPORT_TAB(7) = 93

REPORT_FIELD(8) = "  Waarde"
    REPORT_TAB(8) = 102

REPORT_FIELD(9) = "EV.Eenh."
    REPORT_TAB(9) = 112

REPORT_FIELD(10) = "   Waarde"
    REPORT_TAB(10) = 121

REPORT_TAB(11) = 0

If chkAfdrukInVenster Then
    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = "Inventaris Kontrole"
    Xlog.X.Cols = 12
    Xlog.X.Row = 0
    For T = 0 To 11
        Xlog.X.Col = T
        Xlog.X.Text = REPORT_FIELD(T)
    Next
    Me.Show
End If

End Sub

Private Sub LijstDrukken(InVentaris As Integer)
Dim BeginSleutel    As String
Dim EindSleutel     As String

Dim DCBedrag        As Long
Dim Lijn            As Long

Dim bModDummy       As Integer
Dim XX              As Integer
Dim Flag            As Integer

aa = ""
TLijnen = 0
Lijn = 0
AlgTotBegin = 0
AlgTotAankoop = 0
AlgTotVerkoop = 0
AlgTotEindVoor = 0

Select Case InVentaris
    Case 0
        psTekst(2) = "Artikels te kontroleren " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
    Case 1
        psTekst(2) = "EindInventaris Stock " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
End Select
psTekst(0) = TekstLijn(1).Text
psTekst(3) = TekstLijn(0).Text

InitVelden
bFirst TABLE_PRODUCTS, (Sortering.ListIndex)
If Ktrl Then
    Beep
    Exit Sub
Else
    Screen.MousePointer = vbHourglass
    PAGE_COUNTER = 0
    If chkAfdrukInVenster = 0 Then
        Set Printer = Printers(LijstPrinterNr)
        On Error Resume Next
        Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
        If Printer.Width > 12000 Then
            Printer.FontSize = 10
            Printer.FontName = "Courier New"
            Printer.Print " "
            Printer.FontSize = 10
        Else
            Printer.FontSize = 7.2
            Printer.FontName = "Courier New"
            Printer.Print " "
            Printer.FontSize = 7.2
        End If
    End If
    PrintTitel
    GoSub PrintInfo
End If

Do
    bNext TABLE_PRODUCTS
    If Ktrl Then
        Exit Do
    Else
        GoSub PrintInfo
    End If
Loop
PrintTotaal
Screen.MousePointer = vbNormal
If chkAfdrukInVenster Then
    If chkAfdrukInVenster Then
        Xlog!Kopie(1).Checked = True
        Xlog.X.Row = 1
        Xlog.X.Col = 0
        Xlog.X.ColWidth(0) = 1245
        Xlog.X.ColWidth(1) = 3045
        Xlog.X.ColWidth(2) = 705
        For COUNT_TO = 3 To 10
            Xlog.X.ColWidth(COUNT_TO) = 795
        Next
        Xlog.WindowState = 2
        
        For COUNT_TO = 0 To 2
            Xlog.X.ColAlignment(COUNT_TO) = flexAlignLeftTop
        Next
        For COUNT_TO = 3 To 10
            Xlog.X.ColAlignment(COUNT_TO) = flexAlignRightTop
        Next
        Screen.MousePointer = vbNormal
        Xlog.WijzigenLijn.Visible = False
        Xlog.Afsluiten.TabStop = False
        Xlog.Afsluiten.Enabled = False
        Xlog.cbAfbeelding.Visible = False
        XLogKey = ""
        Xlog.SSTab1.TabVisible(1) = False
        Xlog.Show 1
        Xlog.WindowState = 0
        Unload Xlog
End If
Else
    Printer.EndDoc
End If
Exit Sub

PrintInfo:
RecordToVeld TABLE_PRODUCTS
If vBibTekst(TABLE_PRODUCTS, "#v108 #") = "D" Or vBibTekst(TABLE_PRODUCTS, "#v108 #") = "S" Then
    Return
Else
    VeldTXT(0) = vSet(vBibTekst(TABLE_PRODUCTS, "#v102 #"), 13)  'artikelnummer
    VeldTXT(1) = vSet(vBibTekst(TABLE_PRODUCTS, "#v105 #"), 30)  'naam
    SnelHelpPrint VeldTXT(0), BL_LOGGING
    
    VeldTXT(2) = vSet(vBibTekst(TABLE_PRODUCTS, "#v118 #"), 7) 'rekeningvoorraad
    
    VeldTXT(5) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v119 #")), MASK_SY(7))    'eenheden AK
    VeldTXT(7) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v120 #")), MASK_SY(7))    'eenheden VK
    VeldTXT(3) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v114 #")), MASK_SY(7))    'eenheden BV
        
    If bhEuro Then
        VeldTXT(6) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e121 #")), Right(MASK_SY(4), 8))   'waardeaankoop
        VeldTXT(8) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e122 #")), Right(MASK_SY(4), 8))    'waardeverkoop
        VeldTXT(4) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#e123 #")), Right(MASK_SY(4), 8))    'waardebeginvoorraad
    
        VeldTXT(9) = Dec$(Val(VeldTXT(3)) + Val(VeldTXT(5)) - Val(VeldTXT(7)), MASK_SY(7))  'eindstock
        VeldTXT(10) = Dec$(Val(VeldTXT(4)) + Val(VeldTXT(6)) - Val(VeldTXT(8)), Right(MASK_SY(4), 8))  'eindwaarde
    Else
        VeldTXT(6) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v121 #")), MASK_SY(8))    'waardeaankoop
        VeldTXT(8) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v122 #")), MASK_SY(8))    'waardeverkoop
        VeldTXT(4) = Dec$(Val(vBibTekst(TABLE_PRODUCTS, "#v123 #")), MASK_SY(8))    'waardebeginvoorraad
    
        VeldTXT(9) = Dec$(Val(VeldTXT(3)) + Val(VeldTXT(5)) - Val(VeldTXT(7)), MASK_SY(7))  'eindstock
        VeldTXT(10) = Dec$(Val(VeldTXT(4)) + Val(VeldTXT(6)) - Val(VeldTXT(8)), MASK_SY(8))  'eindwaarde
    End If
    If InVentaris = 1 Then
        Flag = 0
        If cbNegeerRest.Value = 1 Then
            If bhEuro Then
                If Abs(Val(VeldTXT(10))) < Val(TekstLijn(2).Text) Then
                    Flag = 1
                End If
            Else
                If Abs(Val(VeldTXT(10)) / EURO) < Val(TekstLijn(2).Text) Then
                    Flag = 1
                End If
            End If
        End If
        If Flag = 1 Then
        Else
            Flag = 0
            For COUNT_TO = 3 To 8
                If Val(VeldTXT(COUNT_TO)) <> 0 Then
                    Flag = 1
                    Exit For
                End If
            Next
            If Flag = 1 Then
                PrintVelden
            End If
        End If
    ElseIf InVentaris = 0 Then 'kontrolelijstoptie !!!
        If Val(VeldTXT(10)) = 0 And Val(VeldTXT(9)) <> 0 Then  'klopt iets niet
            PrintVelden
        ElseIf Val(VeldTXT(9)) = 0 And Val(VeldTXT(10)) <> 0 Then 'idem
            If cbNegeerRest.Value = 1 Then
                If bhEuro Then
                    If Abs(Val(VeldTXT(10))) < Val(TekstLijn(2).Text) Then
                    Else
                        PrintVelden
                    End If
                Else
                    If Abs(Val(VeldTXT(10)) / EURO) < Val(TekstLijn(2).Text) Then
                    Else
                        PrintVelden
                    End If
                End If
            Else
                PrintVelden
            End If
        ElseIf Val(VeldTXT(9)) < 0 Or Val(VeldTXT(10)) < 0 Then
            PrintVelden
            'For COUNT_TO = 3 To 8
            '    If Val(VeldTXT(COUNT_TO)) < 0 Then
            '        PrintVelden
            '        Exit For
            '    End If
            'Next
        End If
    Else
        PrintVelden
    End If
End If
Lijn = Lijn + 1
Return

DetailCumul:
StartPunt:
bGet TABLE_DUMMY, 0, vSet(vBibTekst(TABLE_JOURNAL, "#v019 #"), 20)
If Ktrl Then
    TLB_RECORD(TABLE_DUMMY) = ""
    vBib TABLE_DUMMY, vBibTekst(TABLE_JOURNAL, "#v019 #"), "v089"
    vBib TABLE_DUMMY, "0", "v013"
    vBib TABLE_DUMMY, "0", "v068"
    bInsert TABLE_DUMMY, 0
    GoTo StartPunt
Else
    RecordToVeld TABLE_DUMMY
    vBib TABLE_DUMMY, Str$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")) + 1), "v013"
    vBib TABLE_DUMMY, Str$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")) + Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))), "v068"
    bUpdate TABLE_DUMMY, 0
End If
Return

End Sub

Private Sub PrintTitel()
Dim T As Integer

If chkAfdrukInVenster Then Exit Sub

    If usrLicentieInfo <> "" Then
        Printer.CurrentX = 50
        Printer.CurrentY = 50
        Printer.Print usrLicentieInfo;
    End If
    PAGE_COUNTER = PAGE_COUNTER + 1
    Printer.CurrentY = 400
    Printer.Print Tab(1); psTekst(2); Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########");

    Printer.Print Tab(108); "Datum  : "; psTekst(0); vbCrLf; vbCrLf;
    Printer.Print Tab(1); UCase$(psTekst(3));

Printer.Print vbCrLf; FULL_LINE
Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print REPORT_FIELD(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop
Printer.Print FULL_LINE; vbCrLf; vbCrLf;

End Sub

Private Sub PrintTotaal()
Dim T As Integer

For T = 0 To 11
    VeldTXT(T) = ""
Next

aa = ""
If chkAfdrukInVenster Then
    aa = aa + Chr(13)
    TLijnen = TLijnen + 1
Else
    Printer.Print vbCrLf; FULL_LINE
End If

If bhEuro Then
    VeldTXT(4) = Dec$(AlgTotBegin, Right(MASK_SY(4), 8))
    VeldTXT(6) = Dec$(AlgTotAankoop, Right(MASK_SY(4), 8))
    VeldTXT(8) = Dec$(AlgTotVerkoop, Right(MASK_SY(4), 8))
    VeldTXT(10) = Dec$(AlgTotEindVoor, Right(MASK_SY(4), 8))
Else
    VeldTXT(4) = Dec$(AlgTotBegin, MASK_SY(8))
    VeldTXT(6) = Dec$(AlgTotAankoop, MASK_SY(8))
    VeldTXT(8) = Dec$(AlgTotVerkoop, MASK_SY(8))
    VeldTXT(10) = Dec$(AlgTotEindVoor, MASK_SY(8))
End If

T = 0
aa = ""
Do While T < 11
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print ; Tab(REPORT_TAB(T));
        Printer.Print ; VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
            'MsgBox "aa = aa + Chr(13)"
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1

End Sub

Private Sub PrintVelden()
Dim T As Integer

aa = ""
T = 0
Do While REPORT_TAB(T) <> 0
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print Tab(REPORT_TAB(T));
        Printer.Print VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
            'MsgBox "aa = aa + Chr(13)"
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
AlgTotBegin = AlgTotBegin + Val(VeldTXT(4))
AlgTotAankoop = AlgTotAankoop + Val(VeldTXT(6))
AlgTotVerkoop = AlgTotVerkoop + Val(VeldTXT(8))
AlgTotEindVoor = AlgTotEindVoor + Val(VeldTXT(10))

If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1: Exit Sub
If Printer.CurrentY >= Printer.Height - 1200 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If

End Sub

Private Sub SubWizard1_GotFocus()

End Sub

Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelLength = Len(TekstLijn(Index).Text)

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 1
        If DATE_INVALID((TekstLijn(1).Text)) Then
            Beep
            TekstLijn(1).Text = MIM_GLOBAL_DATE
            TekstLijn(1).SetFocus
        End If
    Case 3
        TekstLijn(3).Text = Format(Val(TekstLijn(3).Text), "00000")
End Select

End Sub

