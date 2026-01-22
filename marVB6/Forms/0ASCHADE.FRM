VERSION 5.00
Begin VB.Form KTRLSchadeDossiers 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Schadedossiers"
   ClientHeight    =   1890
   ClientLeft      =   2250
   ClientTop       =   2265
   ClientWidth     =   4875
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
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1890
   ScaleWidth      =   4875
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   360
      Left            =   3210
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1410
      Width           =   1560
   End
   Begin VB.OptionButton LijstType 
      Caption         =   "Geopend in aktief boekjaar"
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Value           =   -1  'True
      Width           =   2835
   End
   Begin VB.OptionButton LijstType 
      Caption         =   "Totaallijst afgesloten dossiers"
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2925
   End
   Begin VB.OptionButton LijstType 
      Caption         =   "Lijst dossiers in omloop"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   2805
   End
   Begin VB.OptionButton LijstType 
      Caption         =   "Dossiers voor aktieve klant"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   720
      Width           =   2775
   End
   Begin VB.CommandButton LijstDossiers 
      Caption         =   "Lijst Dossiers"
      Height          =   360
      Left            =   3180
      TabIndex        =   6
      Top             =   540
      Width           =   1590
   End
   Begin VB.CommandButton AfDrukken 
      Caption         =   "Afdruk Dossier"
      Enabled         =   0   'False
      Height          =   390
      Left            =   3210
      TabIndex        =   1
      Top             =   960
      Width           =   1560
   End
   Begin VB.CommandButton ZoekSchade 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "?"
      Height          =   540
      Left            =   2550
      TabIndex        =   0
      Top             =   60
      Width           =   360
   End
   Begin VB.Label LabelInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   240
      Index           =   1
      Left            =   300
      TabIndex        =   9
      Top             =   330
      Width           =   2160
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SchadeNummer"
      Height          =   240
      Index           =   3
      Left            =   330
      TabIndex        =   8
      Top             =   90
      Width           =   1440
   End
End
Attribute VB_Name = "KTRLSchadeDossiers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim fKtrl As Integer
Dim psTekst(3) As String
Dim RapportTitelLijn As String
Dim OmschrijvingsLijn As String

Dim Dossier As String



Private Sub Annuleren_Click()

Unload Me

End Sub


Private Sub Form_Load()

'TeleBibVsoft
GridText = ""

End Sub


Private Sub LijstDossiers_Click()
Dim LijstTyp As Integer
Dim PolisNummer As String * 12
Dim SchadeNummer As String * 12
Dim DatumAfSluiting As String * 8
Dim SchadeDatum As String * 8
Dim CodeNaamBeheerder As String * 15
Dim Verzekeringsnemer As String * 23
Dim TegenPartij As String * 20
Dim VergoedingKlient As Single
Dim VergoedingTegenpartij As Single
Dim Maatschappij As String * 4
Dim ffKtrl As Integer

ffKtrl = 0
RapportTitelLijn = "Dossiernr.   PolisNummer  Mij  Kode/Beheerder  Klient                  Tegenpartij/detail   Schadedt Dt.afslt  Bdrg KL Bdrg TP"

psTekst(2) = "SCHADEDOSSIERS " + Mid(Mim.Caption, InStr(Mim.Caption, "["))

For T = 0 To 3
    If LijstType(T).Value = 0 Then
    Else
        LijstTyp = T
        Exit For
    End If
Next

Select Case LijstTyp
    Case 0
        If LabelInfo(2).Caption = "" Then
            Beep
            Exit Sub
        Else
            OmschrijvingsLijn = "Schades voor klant : " + FVT(TABLE_CUSTOMERS, 0) + " " + FVT(TABLE_CUSTOMERS, 1)
            bGet TABLE_VARIOUS, 0, FVT(TABLE_CUSTOMERS, 0)
            GoTo JumpKlant
        End If
    Case 1
        OmschrijvingsLijn = "Totaallijst openstaande schadegevallen"
    Case 2
        OmschrijvingsLijn = "Totaallijst afgesloten dossiers"
    Case 3
        OmschrijvingsLijn = "Totaallijst dossiers geopend in boekjaar " + Left(BOOKYEAR_FROMTO, 4)
End Select
bFirst TABLE_VARIOUS, 0

JumpKlant:
If Ktrl Then
    Beep
    Exit Sub
Else
    PAGE_COUNTER = 0
    'Printer.Print pPaginaDef; pVerkleind(PrinterType);
    PRINTER_CURRENT_Y = 5
    PrintTitel
    GoSub KontroleVoorPrint
End If

Do
    If ffKtrl Then
        Exit Do
    Else
        bNext TABLE_VARIOUS
    End If
    If Ktrl Then
        Exit Do
    End If
    GoSub KontroleVoorPrint
Loop
Printer.NewPage
Printer.EndDoc
Exit Sub

KontroleVoorPrint:
RecordToVeld TABLE_VARIOUS
If Left(FVT(TABLE_VARIOUS, 1), 1) = "S" Then
Else
    Return
End If

Select Case LijstTyp
    Case 0 'Enkel aktieve klant
        If FVT(TABLE_VARIOUS, 0) <> FVT(TABLE_CUSTOMERS, 0) Then
            ffKtrl = 1
            Return
        End If
    Case 1 'ALLE Openstaande
        If vBibTekst(TABLE_VARIOUS, "#C029 #") <> "" Then
            Return
        End If
    Case 2 'ALLE Afgesloten
        If vBibTekst(TABLE_VARIOUS, "#C029 #") = "" Then
            Return
        End If
    Case 3 'Enkel Aktief boekjaar
        If vBibTekst(TABLE_VARIOUS, "#C200 #") = "" Then
            SchadeDatum = vBibTekst(TABLE_VARIOUS, "#C020 #")
        Else
            SchadeDatum = vBibTekst(TABLE_VARIOUS, "#C200 #")
        End If
        If Right(SchadeDatum, 4) <> Left(BOOKYEAR_FROMTO, 4) Then
            Return
        End If
End Select

SchadeNummer = vBibTekst(TABLE_VARIOUS, "#C000 #")
PolisNummer = vBibTekst(TABLE_VARIOUS, "#A000 #")
If vBibTekst(TABLE_VARIOUS, "#C002 #") = "" Then
    CodeNaamBeheerder = vBibTekst(TABLE_VARIOUS, "#C003 #")
Else
    CodeNaamBeheerder = vBibTekst(TABLE_VARIOUS, "#C002 #") + " " + vBibTekst(TABLE_VARIOUS, "#C003 #")
End If
Maatschappij = Mid(FVT(TABLE_VARIOUS, 1), 2, 4)
If vBibTekst(TABLE_VARIOUS, "#A100 #") = "" Then
    bGet TABLE_CUSTOMERS, 0, FVT(TABLE_VARIOUS, 0)
    If Ktrl Then
        Verzekeringsnemer = ""
    Else
        RecordToVeld TABLE_CUSTOMERS
        Verzekeringsnemer = FVT(TABLE_CUSTOMERS, 1)
    End If
Else
    Verzekeringsnemer = vBibTekst(TABLE_VARIOUS, "#A100 #")
End If
TegenPartij = vBibTekst(TABLE_VARIOUS, "#C100 #")
If vBibTekst(TABLE_VARIOUS, "#C200 #") = "" Then
    SchadeDatum = vBibTekst(TABLE_VARIOUS, "#C020 #")
Else
    SchadeDatum = vBibTekst(TABLE_VARIOUS, "#C200 #")
End If
DatumAfSluiting = vBibTekst(TABLE_VARIOUS, "#C029 #")
VergoedingKlient = Val(vBibTekst(TABLE_VARIOUS, "#C680 #")) + Val(vBibTekst(TABLE_VARIOUS, "#C6B3 #"))
VergoedingTegenpartij = Val(vBibTekst(TABLE_VARIOUS, "#C690 #"))
Printer.Print UCase$(SchadeNummer + " " + PolisNummer + " " + Maatschappij + " " + CodeNaamBeheerder + " " + Verzekeringsnemer + " " + TegenPartij); " "; SchadeDatum; " "; DatumAfSluiting; " "; Dec$((VergoedingKlient), "########"); " "; Dec$(( _
VergoedingTegenpartij), "########");
PRINTER_CURRENT_Y = PRINTER_CURRENT_Y + 1
Printer.Print vbCrLf;
If Printer.CurrentY >= Printer.Height - 1200 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If
Return

End Sub

Private Sub NieuweSchadeInvoer()

End Sub


Private Sub PrintTitel()
Dim T As Integer

  If usrLicentieInfo <> "" Then
        Printer.CurrentX = 50
        Printer.CurrentY = 50
        Printer.Print usrLicentieInfo;
    End If
    PAGE_COUNTER = PAGE_COUNTER + 1
    Printer.CurrentY = 400
    PRINTER_CURRENT_Y = 3
    Printer.Print Tab(1); psTekst(2); Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########");
    Printer.Print Tab(108); "Datum  : "; Format(Now, "dd/mm/yyyy"); vbCrLf; vbCrLf;
    Printer.Print Tab(1); OmschrijvingsLijn;

Printer.Print vbCrLf$; FULL_LINE; vbCrLf;
PRINTER_CURRENT_Y = PRINTER_CURRENT_Y + 1

Printer.Print RapportTitelLijn
Printer.Print FULL_LINE; vbCrLf; vbCrLf;
PRINTER_CURRENT_Y = PRINTER_CURRENT_Y + 2

End Sub

Private Sub SchadeKiezen(Zoekstring As String)
ReDim TeleBibString(50) As String * 1018
Dim SchadeNummer As String * 17
Dim SchadeDatum As String * 8
Dim SchadeOmschrijving As String * 30
Dim SchadeReferte As String * 12

Xlog.Left = 0
Xlog.Width = Screen.Width
Xlog.X.Width = Screen.Width - 2
'Xlog.InfoLabel.Caption = "Schadenr.          Datum      Tegenpartij/omschrijving       Onze Referte"
Xlog.Caption = "Schadedossiers : " + RTrim$(FVT(TABLE_CUSTOMERS, 1)) + " " + RTrim$(FVT(TABLE_CUSTOMERS, 2))

'Do While Xlog.X.ListCount
'    Xlog.X.RemoveItem 0
'Loop

SchadeNummer = FVT(TABLE_VARIOUS, 1)
SchadeDatum = vBibTekst(TABLE_VARIOUS, "#C020 #")
SchadeOmschrijving = vBibTekst(TABLE_VARIOUS, "#C100 #")
SchadeReferte = vBibTekst(TABLE_VARIOUS, "#C010 #")
Xlog.X.AddItem SchadeNummer + Chr$(179) + SchadeDatum + Chr$(179) + SchadeOmschrijving + Chr$(179) + SchadeReferte
Do
    bNext TABLE_VARIOUS
    If Ktrl Or KEY_BUF(TABLE_VARIOUS) <> Zoekstring Then
        Exit Do
    Else
        RecordToVeld TABLE_VARIOUS
        If Left(FVT(TABLE_VARIOUS, 1), 1) <> "S" Then
        Else
            SchadeNummer = FVT(TABLE_VARIOUS, 1)
            SchadeDatum = vBibTekst(TABLE_VARIOUS, "#C020 #")
            SchadeOmschrijving = vBibTekst(TABLE_VARIOUS, "#C100 #")
            SchadeReferte = vBibTekst(TABLE_VARIOUS, "#C010 #")
            Xlog.X.AddItem SchadeNummer + Chr$(179) + SchadeDatum + Chr$(179) + SchadeOmschrijving + Chr$(179) + SchadeReferte
        End If
    End If
Loop
Xlog.X.AddItem "Dossier +        " + Chr$(179) + "00000000" + Chr$(179)
Screen.MousePointer = vbNormal
ShowThis:
'Xlog.X.ListIndex = 0
Xlog.SSTab1.TabVisible(1) = False
Xlog.Show 1

If XLogKey = "" Then
    Exit Sub
ElseIf Left(XLogKey, 12) = "Dossier +   " Then
    Msg$ = "Nieuw schadedossier bijvoegen !  Bent U Zeker ?"
    Ktrl = MsgBox(Msg$, 292)
    If Ktrl = 6 Then
        GridText = LabelInfo(2).Caption + Chr$(13)
        fKtrl = 99
        NieuweSchadeInvoer
        Exit Sub
    Else
        Exit Sub
    End If
Else
    FVT(TABLE_VARIOUS, 0) = ""
    FVT(TABLE_VARIOUS, 1) = Left(XLogKey, InStr(XLogKey, vbCrLf) - 1)
    bGet TABLE_VARIOUS, 1, FVT(TABLE_VARIOUS, 1)
    If Ktrl Then
        LabelInfo(1).Caption = ""
        Beep
        Exit Sub
    Else
        RecordToVeld TABLE_VARIOUS
        LabelInfo(1).Caption = FVT(TABLE_VARIOUS, 1)
        GridText = ""
        T = 0
        Do While TELEBIB_CODE(T) <> Space$(10)
            GridText = GridText + vBibTekst(TABLE_VARIOUS, "#" + Mid(TELEBIB_CODE(T), 5, 5) + "#") + Chr$(13)
            T = T + 1
        Loop
        fKtrl = 100
        NieuweSchadeInvoer
        Exit Sub
    End If
End If

End Sub

Private Sub SchrijfWeg()
Dim FlNetTemp As Integer
On Local Error GoTo SchrijfWegError

FlNetTemp = FreeFile
Open LOCATION_ASWEB + "AS1.SND" For Append As FlNetTemp
Print #FlNetTemp, A; vbCrLf;
Close FlNetTemp
    Exit Sub

SchrijfWegError:
    MsgBox "Openen van " + LOCATION_ASWEB + "AS1.SND onmogelijk.  Probeer opnieuw"
    Resume

End Sub



Private Sub ZoekSchade_Click()
Dim ZoekIndex As Integer
Dim Zoekstring As String

    bGet TABLE_CONTRACTS, 0, Venster!TxtInfo(0).Text
    If Ktrl Then
        MsgBox "Waw!"
        Exit Sub
    Else
        RecordToVeld TABLE_CONTRACTS
        bGet TABLE_SUPPLIERS, 0, "CO" + vBibTekst(TABLE_CONTRACTS, "#A010 #")
        ktrlHier = Ktrl
        bGet TABLE_CUSTOMERS, 0, vBibTekst(TABLE_CONTRACTS, "#A100 #")
        ktrlHier = ktrlHier + Ktrl
        If ktrlHier Then
            Beep
            LabelInfo(2).Caption = ""
            Exit Sub
        Else
            RecordToVeld TABLE_SUPPLIERS
            RecordToVeld TABLE_CUSTOMERS
        End If
    End If

'vBib TABLE_VARIOUS, vSet("K" + vBibTekst(TABLE_CONTRACTS, "#A110 #"), 13), "v004"
'vBib TABLE_VARIOUS, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12), "A110"
'vBib TABLE_VARIOUS, Mid(a, 58, 4), "A010"
'vBib TABLE_VARIOUS, PolisNummer, "A000"
'vBib TABLE_VARIOUS, Dossier, "C000"
'vBib TABLE_VARIOUS, vSet("24" + Mid(a, 58, 4) + Dossier, 20), "v005"

bGet TABLE_VARIOUS, ZoekIndex, "24" + vBibTekst(TABLE_CONTRACTS, "#A010 #") + Dossier
If Ktrl And ZoekIndex = 0 Then
    Msg$ = "Een nieuw schadedossier ?"
    KtrlBox = MsgBox(Msg$, vbYesNo)
    If KtrlBox = vbYes Then
        NieuweSchadeInvoer
    Else
        Exit Sub
    End If
ElseIf Ktrl And ZoekIndex = 1 Then
    MsgBox "Dit schadenummer bestaat niet." + vbCrLf + vbCrLf + "Kontroleer, of zoek via polis"
    LabelInfo(1).Caption = ""
    Exit Sub
Else
    bGetOrGreater TABLE_VARIOUS, ZoekIndex, Zoekstring
    RecordToVeld TABLE_VARIOUS
    bGet TABLE_CUSTOMERS, 0, FVT(TABLE_VARIOUS, 0)
    RecordToVeld TABLE_CUSTOMERS
    SchadeKiezen Zoekstring
End If

End Sub

Private Sub ZoekSchade_LostFocus()

ZoekSchade.Default = False

End Sub

