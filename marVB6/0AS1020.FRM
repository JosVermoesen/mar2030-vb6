VERSION 5.00
Begin VB.Form AS1VrijBericht 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vrije Berichtgeving MAATSCHAPPIJEN"
   ClientHeight    =   3855
   ClientLeft      =   495
   ClientTop       =   780
   ClientWidth     =   8895
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
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   8895
   Begin VB.ComboBox KeuzeInfo 
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
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   30
      Width           =   4545
   End
   Begin VB.ComboBox KeuzeInfo 
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
      Height          =   315
      Index           =   1
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   690
      Width           =   3705
   End
   Begin VB.ComboBox KeuzeInfo 
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
      Height          =   315
      Index           =   5
      Left            =   4740
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   300
      Width           =   3105
   End
   Begin VB.ComboBox KeuzeInfo 
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
      Height          =   315
      Index           =   2
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1020
      Width           =   3345
   End
   Begin VB.ComboBox KeuzeInfo 
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
      Height          =   315
      Index           =   4
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1350
      Width           =   3345
   End
   Begin VB.ComboBox KeuzeInfo 
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
      Height          =   315
      Index           =   3
      Left            =   4740
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   630
      Width           =   3090
   End
   Begin VB.ComboBox KeuzeInfo 
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
      Height          =   315
      Index           =   6
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   360
      Width           =   4545
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   4710
      TabIndex        =   11
      Top             =   1380
      Width           =   2520
   End
   Begin VB.TextBox TekstInfo 
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
      Index           =   0
      Left            =   1320
      TabIndex        =   13
      Top             =   1680
      Width           =   3345
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   5850
      TabIndex        =   16
      Top             =   1020
      Width           =   2040
   End
   Begin VB.CommandButton Wegschrijven 
      Caption         =   "AS1.SND bijwerken"
      Height          =   630
      Left            =   7470
      TabIndex        =   18
      Top             =   1380
      Width           =   1395
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
      Left            =   8040
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton CmdZoekPolis 
      Caption         =   "Polis &Zoeken"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   570
      Left            =   7920
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   420
      Width           =   915
   End
   Begin VB.CheckBox AlleMijen 
      Caption         =   "Alle ASSURNET mijen"
      Height          =   240
      Left            =   5580
      TabIndex        =   19
      Top             =   30
      Width           =   2250
   End
   Begin VB.TextBox VrijeTekst 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1725
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   2040
      Width           =   8800
   End
   Begin VB.CheckBox AfdrukBericht 
      Caption         =   "&Briefwisseling i.p.v. AS1"
      Height          =   240
      Left            =   4740
      TabIndex        =   17
      Top             =   1710
      Width           =   2460
   End
   Begin VB.Label Label1 
      Caption         =   "&Dienst"
      Height          =   210
      Index           =   6
      Left            =   150
      TabIndex        =   4
      Top             =   720
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "&Tak"
      Height          =   240
      Index           =   7
      Left            =   150
      TabIndex        =   7
      Top             =   1020
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "&Onderwerp"
      Height          =   240
      Index           =   10
      Left            =   150
      TabIndex        =   12
      Top             =   1710
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "&Referentie"
      Height          =   240
      Index           =   11
      Left            =   150
      TabIndex        =   9
      Top             =   1380
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "&Zetel"
      Height          =   240
      Index           =   13
      Left            =   4800
      TabIndex        =   0
      Top             =   30
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Ter &attentie"
      Height          =   240
      Index           =   12
      Left            =   4740
      TabIndex        =   15
      Top             =   1050
      Width           =   1050
   End
End
Attribute VB_Name = "AS1VrijBericht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim a As String * 140
Dim aa As String
Dim FlAsNet As Integer
Dim MaxLijn79 As Integer
Dim KontaktPersoon As Integer

Dim VsoftVanaf          As String
Dim VsoftTot            As String
Dim BeginXbox           As Integer
Dim BeginYbox           As Integer

Private Sub AfdrukBericht_Click()

MaxLijn79 = 79
If AfdrukBericht.Value = 0 Then
Else
    Wegschrijven.Caption = "Afdrukken"
    AfdrukBericht.Enabled = False
End If

End Sub

Private Sub AlleMijen_Click()

Select Case AlleMijen.Value
    Case 0
        KeuzeInfo(0).Visible = True
    Case 1
        KeuzeInfo(4).ListIndex = 2
        KeuzeInfo(0).Visible = False
    Case Else
        MsgBox "STOP"
End Select

End Sub

Private Sub DrukAf()
ReDim rSip(5) As String
Dim Taal As Integer
Dim FlTemp As Integer
Dim T As Integer
Dim Tekst As String
Dim TabIs As Integer
Dim Msg As String

bGet FlLeverancier, 0, vSet(Left$(KeuzeInfo(6).Text, 2) + Left$(KeuzeInfo(0).Text, 4), 12)
If Ktrl Then
    MsgBox "stop"
Else
    RecordToVeld FlLeverancier
End If

PrintSteedsOpnieuw:
Tekst = VrijeTekst.Text
' If AfDrukBericht.Value = 1 Then
    TabIs = 10
'Else
'    TabIs = 20
'End If

'On Local Error GoTo FileHandler

Taal = Val(vBibTekst(FlLeverancier, "#A10C #"))
If Val(vBibTekst(FlLeverancier, "#A102 #")) = 0 Then
    rSip(0) = vBibTekst(FlLeverancier, "#A100 #")
Else
    rSip(0) = Mid$(fmarBoxText("003", Format$(Taal), vBibTekst(FlLeverancier, "#A102 #")), 4, 10) + " " + vBibTekst(FlLeverancier, "#A100 #")
End If

If Val(vBibTekst(FlLeverancier, "#vs01 #")) = 0 Then
    rSip(1) = vBibTekst(FlLeverancier, "#A125 #")
Else
    rSip(1) = Mid$(fmarBoxText("003", Format$(Taal), vBibTekst(FlLeverancier, "#vs01 #")), 4, 10) + " " + vBibTekst(FlLeverancier, "#A125 #")
End If

rSip(2) = vBibTekst(FlLeverancier, "#A104 #") + " " + vBibTekst(FlLeverancier, "#A105 #") + " " + vBibTekst(FlLeverancier, "#A106 #")
rSip(4) = vBibTekst(FlLeverancier, "#A109 #") + " " + vBibTekst(FlLeverancier, "#A107 #") + " " + vBibTekst(FlLeverancier, "#A108 #")

'On Local Error GoTo PrintHandler

Page = 0
    GoSub KopBalk
    GoSub DetailLijnen
Printer.EndDoc
AS1VrijBericht.SetFocus
      Msg = "Herdrukken van dit dokument ?" + vbCrLf
Msg = Msg + "Bent U zeker ?" + vbCrLf + vbCrLf
Msg = Msg + "Cancel of ESC om nog iets te wijzigen..."
Ktrl = MsgBox(Msg$, 3, "Afdruk via printer")
Select Case Ktrl
    Case 2, 7
    Case Else
        GoTo PrintSteedsOpnieuw
End Select
Exit Sub

KopBalk:
PrintUserDef "1" + Format$(Taal)
Printer.CurrentX = BeginXbox + 350
Printer.CurrentY = BeginYbox + 350
Printer.FontSize = 10
Printer.FontBold = True
For tSip = 0 To 4
    Printer.Print UCase$(rSip(tSip)); vbCrLf;
    Printer.CurrentX = BeginXbox + 350
Next

Printer.CurrentY = Val(VsoftVanaf)
Printer.CurrentX = 0
Printer.FontSize = 10

Printer.Print Tab(TabIs);
TabBold = Printer.CurrentX

Printer.Print Tab(60); UCase(Format(Now, "DDDD dd/mm/yyyy hh:mm")); vbCrLf; vbCrLf;
Printer.Print vbCrLf;

pfTl = 0
Page = Page + 1
If Left$(KeuzeInfo(6).Text, 2) = "CO" Then
    Printer.Print Tab(10); "AGENT  : "; Right$(KeuzeInfo(0).Text, 8); vbCrLf; vbCrLf;
End If

Printer.Print Tab(10); "T.A.V. : ";
Printer.Print TekstInfo(2).Text;
Printer.Print vbCrLf;

If Left$(KeuzeInfo(6).Text, 2) = "CO" Then
    Printer.Print Tab(10); "ZETEL  : ";
    Printer.FontBold = True
    Printer.Print Right$(KeuzeInfo(5).Text, Len(KeuzeInfo(5).Text) - 3);
    Printer.Print vbCrLf;

    Printer.Print Tab(10); "DIENST : ";
    Printer.Print Right(KeuzeInfo(1).Text, Len(KeuzeInfo(1).Text) - 3);
    Printer.Print vbCrLf;
    
    Printer.Print Tab(10); "TAK    : ";
    Printer.Print Right(KeuzeInfo(2).Text, Len(KeuzeInfo(2).Text) - 3); " "; Right$(KeuzeInfo(3).Text, Len(KeuzeInfo(3).Text) - 3);
    Printer.Print vbCrLf;
End If

Printer.Print Tab(10); "INZAKE : ";
Printer.Print TekstInfo(0).Text;
Printer.Print vbCrLf;

If Left$(KeuzeInfo(6).Text, 2) = "CO" Then
    Printer.Print Tab(10); Right$(KeuzeInfo(4).Text, Len(KeuzeInfo(4).Text) - 3) + " : ";
    Printer.Print TekstInfo(1).Text;
    Printer.Print vbCrLf;
End If

Printer.FontName = "Courier New"
Printer.FontBold = False
Printer.FontSize = 10
Printer.Print vbCrLf; vbCrLf;
pfTl = 0
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
    pfTl = pfTl + 1
    If Printer.CurrentY >= Printer.Height - 1200 Then
        Printer.NewPage
        Printer.CurrentY = 400
        If AfdrukBericht.Value = 0 Then
            Printer.Print vbCrLf; vbCrLf; Tab(TabIs); "** Kontroleblad Bericht via ASSURNET.  Tijdstip opmaak : "; Format$(Now, "hh:mm:ss"); " **"; vbCrLf;
        End If
    End If
Loop
If AfdrukBericht.Value = 0 Then
    Printer.Print vbCrLf; vbCrLf; Tab(TabIs); "** Kontroleblad Bericht via ASSURNET.  Tijdstip opmaak : "; Format$(Now, "hh:mm:ss"); " **"; vbCrLf;
End If
Printer.NewPage
Return

PrintHandler:
MsgBox "Stopkode extern toestel." + vbCrLf + vbCrLf + "Kontroleer, of rapporteer aan Vsoft.", 0, "Onverwachte situatie"
Resume

End Sub

Private Sub Form_Load()
Top = 0
Left = 0

PaginaTeller = 0
KeuzeInfo(1).AddItem "A: Produktie/uitgifte"
KeuzeInfo(1).AddItem "B: Boekhouding"
KeuzeInfo(1).AddItem "C: Schade"
KeuzeInfo(1).AddItem "D: Commercieel"
KeuzeInfo(1).AddItem "E: Informatica en Organisatie"
KeuzeInfo(1).AddItem "H: Herverzekering"
KeuzeInfo(1).ListIndex = 0

KeuzeInfo(2).AddItem "1: Leven"
KeuzeInfo(2).AddItem "2: Ziekte en individuele"
KeuzeInfo(2).AddItem "3: Brand - gewone risiko's"
KeuzeInfo(2).AddItem "4: Burgerlijke aansprakelijkheid"
KeuzeInfo(2).AddItem "5: Auto"
KeuzeInfo(2).AddItem "6: Arbeidsongevallen"
KeuzeInfo(2).AddItem "7: Brand - industriele risiko's"
KeuzeInfo(2).AddItem "8: Transport"
KeuzeInfo(2).AddItem "9: Diverse (Rechtsbijstand e.a.)"
KeuzeInfo(2).AddItem "0: Leningen"
KeuzeInfo(2).ListIndex = 8

KeuzeInfo(3).AddItem "O: Ondernemingen"
KeuzeInfo(3).AddItem "P: Partikulieren"
KeuzeInfo(3).AddItem " : Onbepaald"
KeuzeInfo(3).ListIndex = 1

KeuzeInfo(4).AddItem "P: Polis"
KeuzeInfo(4).AddItem "S: Schadegeval"
KeuzeInfo(4).AddItem " : Andere"
KeuzeInfo(4).ListIndex = 2

KeuzeInfo(5).AddItem "10: Brussel"
KeuzeInfo(5).AddItem "20: Antwerpen"
KeuzeInfo(5).AddItem "23: Turnhout"
KeuzeInfo(5).AddItem "27: Sint-Niklaas"
KeuzeInfo(5).AddItem "28: Mechelen"
KeuzeInfo(5).AddItem "30: Leuven"
KeuzeInfo(5).AddItem "35: Hasselt"
KeuzeInfo(5).AddItem "40: Liege"
KeuzeInfo(5).AddItem "48: Verviers"
KeuzeInfo(5).AddItem "49: Eupen"
KeuzeInfo(5).AddItem "50: Namur"
KeuzeInfo(5).AddItem "60: Charleroi"
KeuzeInfo(5).AddItem "67: Arlon"
KeuzeInfo(5).AddItem "70: Mons"
KeuzeInfo(5).AddItem "75: Tournai"
KeuzeInfo(5).AddItem "80: Brugge"
KeuzeInfo(5).AddItem "84: Oostende"
KeuzeInfo(5).AddItem "85: Kortrijk"
KeuzeInfo(5).AddItem "90: Gent"
KeuzeInfo(5).AddItem "93: Aalst"
KeuzeInfo(5).ListIndex = 0

KeuzeInfo(6).AddItem "CO: Verzekeringsmaatschappijen"
KeuzeInfo(6).AddItem "EX: Experten"
KeuzeInfo(6).AddItem "IN: Inspecteurs"
KeuzeInfo(6).AddItem "AD: Advokaten"
KeuzeInfo(6).ListIndex = 0

InstallMij Left$(KeuzeInfo(6).Text, 2)
MaxLijn79 = 79

End Sub

Private Sub Form_Unload(Cancel As Integer)

bClose FlAllerlei

End Sub

Private Sub InstallMij(VoorLetter As String)
Dim T As Integer

KeuzeInfo(0).Clear
T = -1
bGetOrGreater FlLeverancier, 0, vSet(VoorLetter, 12)
If Ktrl Then
    Beep
    Exit Sub
Else
    RecordToVeld FlLeverancier
End If
Screen.MousePointer = 11
If Left$(KeyBuf(FlLeverancier), 2) <> VoorLetter Then
    Beep
Else
    T = T + 1
    KeuzeInfo(0).AddItem Mid$(vBibTekst(FlLeverancier, "#A110 #"), 3, 4) + ": " + vBibTekst(FlLeverancier, "#A100 #") + "/" + vSet(vBibTekst(FlLeverancier, "#A400 #"), 8)
    Do
        bNext FlLeverancier
        If Ktrl Or Left$(KeyBuf(FlLeverancier), 2) <> VoorLetter Then
            Exit Do
        Else
            RecordToVeld FlLeverancier
            T = T + 1
            KeuzeInfo(0).AddItem Mid$(vBibTekst(FlLeverancier, "#A110 #"), 3, 4) + ": " + vBibTekst(FlLeverancier, "#A100 #") + "/" + vSet(vBibTekst(FlLeverancier, "#A400 #"), 8)
        End If
    Loop
End If
Screen.MousePointer = 0
If KeuzeInfo(0).ListCount Then
    KeuzeInfo(0).ListIndex = 0
End If

End Sub

Private Sub KeuzeInfo_Click(Index As Integer)

Select Case Index
    Case 0
        'MsgBox "nog in te voegen"
        If KTRLVerzoek("020", Left$(KeuzeInfo(0).Text, 4)) Then
            Wegschrijven.Caption = "AS1.SND bijwerken"
            AfdrukBericht.Enabled = True
            AfdrukBericht.Value = 0
        Else
            Wegschrijven.Caption = "Afdrukken"
            AfdrukBericht.Enabled = False
            AfdrukBericht.Value = 1
        End If
    Case 4
        Select Case Left(KeuzeInfo(4), 1)
            Case "P"
                CmdZoekPolis.Enabled = True
                'cmdzoekschade.enabled = False
            Case "S"
                CmdZoekPolis.Enabled = False
                'cmdzoekschade.enabled = true
            Case Else
                CmdZoekPolis.Enabled = False
                'cmdzoekschade.enabled = False
        End Select
            
    Case 6
        InstallMij Left$(KeuzeInfo(6).Text, 2)
        Do While KeuzeInfo(1).ListCount
            KeuzeInfo(1).RemoveItem 0
        Loop
        Select Case Left$(KeuzeInfo(6).Text, 2)
            Case "CO"
                KeuzeInfo(1).AddItem "A: Produktie/uitgifte"
                KeuzeInfo(1).AddItem "B: Boekhouding"
                KeuzeInfo(1).AddItem "C: Schade"
                KeuzeInfo(1).AddItem "D: Commercieel"
                KeuzeInfo(1).AddItem "E: Informatica en Organisatie"
                KeuzeInfo(1).AddItem "H: Herverzekering"
                KeuzeInfo(1).ListIndex = 0
                For T = 2 To 5
                    KeuzeInfo(T).Visible = True
                Next
                Label1(7).Visible = True
                Label1(11).Visible = True
                Label1(13).Visible = True
                AlleMijen.Visible = True
            Case Else
                KeuzeInfo(1).AddItem "D: Diverse"
                KeuzeInfo(1).AddItem "A: Aankoopdienst"
                KeuzeInfo(1).AddItem "B: Boekhouding"
                KeuzeInfo(1).AddItem "C: Commercieel"
                KeuzeInfo(1).AddItem "I: Informatica en Organisatie"
                KeuzeInfo(1).AddItem "T: Technische dienst"
                KeuzeInfo(1).AddItem "V: Verkoopdienst"
                KeuzeInfo(1).ListIndex = 0
                For T = 2 To 5
                    KeuzeInfo(T).Visible = False
                Next
                Label1(7).Visible = False
                Label1(11).Visible = False
                Label1(13).Visible = False
                AlleMijen.Visible = False
        End Select
End Select

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

If Dir$(BedrijfsLokatie + "Dok" + TypeEnTaal + ".Txt") = "" Then
    Beep
    Exit Sub
Else
    FlFree = FreeFile
    Open BedrijfsLokatie + "Dok" + TypeEnTaal + ".Txt" For Input As FlFree
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
Exit Sub

ErrorLoad:
MsgBox "Stop tijdens inladen dokumentdefinitie" + vbCrLf + vbCrLf + Error$
Exit Sub
Resume

End Sub

Private Sub Schoon()

KeuzeInfo(1).ListIndex = 0
KeuzeInfo(2).ListIndex = 4
KeuzeInfo(3).ListIndex = 1
KeuzeInfo(4).ListIndex = 0
VrijeTekst.Text = ""
For T = 0 To 2
    TekstInfo(T).Text = ""
Next

End Sub

Private Sub SchrijfWeg(RecordLijn As String)

On Local Error GoTo SchrijfWegError

Print #FlAsNet, RecordLijn
If Mid$(RecordLijn, 50, 1) = "X" Then
    Close FlAsNet
End If
Exit Sub

SchrijfWegError:
    FlAsNet = FreeFile
    Open AssurnetLokatie + "AS1.SND" For Append Lock Write As FlAsNet
    Resume

End Sub

Private Sub sluiten_Click()

Unload AS1VrijBericht

End Sub

Private Sub VoorVoegsel()

a = ""
Mid$(a, 1, 3) = "AS1"                   'constante
Mid$(a, 16, 1) = "0"                    'constante
Mid$(a, 17, 3) = "020"                  'vrij bericht
Mid$(a, 20, 1) = "0"                    'vrije inhoud
Mid$(a, 21, 1) = "1"                    'producent
Mid$(a, 22, 8) = vSet(ProducentNummer, 8)    'ASNET prod.nr.
Mid$(a, 30, 12) = Format$(Now, "yymmddhhmmss") 'datum
Mid$(a, 42, 1) = "1"                    'gewoon niet vertrouwelijk
If Left$(KeuzeInfo(0).Text, 4) = "9000" Then
    Mid$(a, 43, 1) = "4"                    'bestemmeling=assurnet
Else
    Mid$(a, 43, 1) = "2"                    'bestemmeling=mij.
End If
Mid$(a, 44, 4) = Left$(KeuzeInfo(0).Text, 4) 'BVVO kode
Mid$(a, 48, 2) = "00"                   'volgnummer
Mid$(a, 50, 1) = "X"                    'vervolg tekst
Mid$(a, 51, 40) = vSet((TekstInfo(0).Text), 40) 'titel
Mid$(a, 91, 1) = Left$(KeuzeInfo(4).Text, 1)    'Referentietype
Mid$(a, 92, 12) = vSet((TekstInfo(1).Text), 12) 'Referentie nr
Mid$(a, 104, 1) = Left$(KeuzeInfo(1).Text, 1) 'dienstkode
Mid$(a, 105, 1) = Left$(KeuzeInfo(2).Text, 1) 'takkode
Mid$(a, 106, 1) = Left$(KeuzeInfo(3).Text, 1) 'klientkode
Mid$(a, 107, 10) = vSet((TekstInfo(2).Text), 10) 'bijkomende identificatie mij.
Mid$(a, 117, 2) = Left$(KeuzeInfo(5).Text, 2) 'kode regionale zetel

If AlleMijen.Value = 1 Then
    Mid$(a, 43, 1) = "3"
    Mid$(a, 44, 4) = "3LR "
End If

End Sub

Private Sub VrijeTekst_GotFocus()

Select Case Left(KeuzeInfo(4), 1)
    Case "P"
        CmdZoekPolis.Default = False
    Case "S"
        'cmdzoekschade.Default = False
End Select

End Sub

Private Sub VrijeTekst_LostFocus()

Select Case Left(KeuzeInfo(4), 1)
    Case "P"
        CmdZoekPolis.Default = True
    Case "S"
        'cmdzoekschade.Default = true
End Select


End Sub

Private Sub Wegschrijven_Click()
Dim TelLijn         As Long
Dim positieCRLF     As Long
Dim SpatieLijn      As Long
Dim Tekst           As String
Dim TekstLijn       As String
Dim Msg             As String
Dim MijKeuze        As String

If VrijeTekst.Text = "" Then
    Exit Sub
Else
    VoorVoegsel
    TelLijn = 0
    Tekst = VrijeTekst.Text
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

If AlleMijen.Value Then
    MijKeuze = "ALLE ASSURNET maatschappijen !!!!"
Else
    MijKeuze = KeuzeInfo(0).Text
End If

If AfdrukBericht.Value = 0 Then
          Msg = "Afdruk kontroleblad én versturen via ASSURNET van " + vbCrLf
    Msg = Msg + Format$(TelLijn) + " tekstlijnen voor " + MijKeuze + vbCrLf + vbCrLf
    Msg = Msg + "Bent U zeker ?" + vbCrLf + vbCrLf
    Msg = Msg + "Kies No om enkel te versturen via ASSURNET !" + vbCrLf
    Msg = Msg + "Cancel of ESC om terug te keren..."
    Ktrl = MsgBox(Msg$, 3, "Vrij bericht versturen en/of afdrukken")
    Select Case Ktrl
        Case 2
            Exit Sub
        Case 7
        Case Else
            DrukAf
            If Ktrl = 2 Then
                Exit Sub
            End If
    End Select

    TelLijn = 0
    aa = ""
    Tekst = VrijeTekst.Text
    Do While Tekst <> ""
        If TelLijn = 14 Then
            Mid$(a, 48, 2) = Format$(Val(Mid$(a, 48, 2)) + 1, "00")
            Mid$(a, 50, 1) = "S"
            SchrijfWeg a + aa
            aa = ""
            TelLijn = 0
        End If
        TelLijn = TelLijn + 1
        positieCRLF = InStr(Tekst, vbCrLf)
        If positieCRLF Then
            aa = aa + vSet(Left$(Tekst, positieCRLF - 1), 79)
            Tekst = Right$(Tekst, Len(Tekst) - positieCRLF - 1)
        Else
            aa = aa + vSet(Tekst, 79)
            Tekst = ""
        End If
    Loop
    For SpatieLijn = TelLijn + 1 To 14
        aa = aa + Space$(79)
    Next
    Mid$(a, 48, 2) = Format$(Val(Mid$(a, 48, 2)) + 1, "00")
    Mid$(a, 50, 1) = "X"
    SchrijfWeg a + aa
    Schoon
Else
          Msg = "Afdrukken van " + Str$(TelLijn) + " tekstlijnen" + vbCrLf
    Msg = Msg + "voor " + MijKeuze + vbCrLf + vbCrLf
    Msg = Msg + "Bent U zeker ?"
    Ktrl = MsgBox(Msg$, 292, "Afdruk via PRINTER")
    If Ktrl = 6 Then
        DrukAf
        If Ktrl = 2 Then
            Exit Sub
        Else
            Schoon
        End If
    End If
End If

End Sub

Private Sub CmdZoekPolis_Click()

KeuzeInfo(4).ListIndex = 0
Venster.Hide
Venster.Show 1
AS1VrijBericht.SetFocus
TekstInfo(1).Text = vBibTekst(FlPolis, "#A000 #")
KeuzeInfo(0).ListIndex = 0
Do While Left$(KeuzeInfo(0).Text, 4) <> "9000"
    If Left$(KeuzeInfo(0).Text, 4) = vBibTekst(FlPolis, "#A010 #") Then
        Exit Do
    Else
        KeuzeInfo(0).ListIndex = KeuzeInfo(0).ListIndex + 1
    End If
Loop
KeuzeInfo(1).ListIndex = 0
KeuzeInfo(2).ListIndex = 0
Do While KeuzeInfo(2).ListIndex <> 9
    If Left$(KeuzeInfo(2).Text, 1) = Left$(vBibTekst(FlPolis, "#vs96 #"), 1) Then
        Exit Do
    Else
        KeuzeInfo(2).ListIndex = KeuzeInfo(2).ListIndex + 1
    End If
Loop
If Left$(KeuzeInfo(2).Text, 1) <> Left$(vBibTekst(FlPolis, "#vs96 #"), 1) Then
    MsgBox "Takkode van deze polis onbekend.  Verbeter a.u.b. !"
End If
bGet FlAllerlei, 1, vSet("25" + vBibTekst(FlPolis, "#A010 #") + vBibTekst(FlPolis, "#A000 #"), 20)
If Ktrl Then
    'niks
Else
    RecordToVeld FlAllerlei
    Select Case Left(vBibTekst(FlAllerlei, "#A500 #"), 1)
        Case "5"    'Voertuigen
            GridText = "Klient vervangt zijn voertuig door een ander met volgende kenmerken :" + vbCrLf + vbCrLf
            GridText = GridText + "VOERTUIG Aard : " + fmarBoxText("21", "2", vBibTekst(FlAllerlei, "#5003 #")) + vbCrLf
            GridText = GridText + "Merk : " + vBibTekst(FlAllerlei, "#500A #") + vbCrLf
            GridText = GridText + "Type : " + vBibTekst(FlAllerlei, "#500B #") + vbCrLf
            GridText = GridText + "P.V.G. nummer : " + vBibTekst(FlAllerlei, "#5004 #") + vbCrLf
            GridText = GridText + "Bouwjaar : " + vBibTekst(FlAllerlei, "#5002 #") + vbCrLf
            GridText = GridText + "Plaatnummer : " + vBibTekst(FlAllerlei, "#5020 #") + vbCrLf
            GridText = GridText + "Onderstelnummer : " + vBibTekst(FlAllerlei, "#5021 #") + vbCrLf
            GridText = GridText + "Datum 1e omloop : " + vBibTekst(FlAllerlei, "#5060 #") + vbCrLf
            GridText = GridText + "Vermogen of cylinderinhoud : " + vBibTekst(FlAllerlei, "#5011 #") + vbCrLf
            GridText = GridText + "Eenheid vermogen : " + fmarBoxText("22", "2", vBibTekst(FlAllerlei, "#5010 #")) + vbCrLf
            GridText = GridText + "Brandstof : " + fmarBoxText("23", "2", vBibTekst(FlAllerlei, "#5015 #")) + vbCrLf + vbCrLf
            GridText = GridText + "De aanvraag tot inschrijving is rechtstreeks overgemaakt aan D.I.V."
            KtrlBox = MsgBox(GridText, vbYesNo + vbDefaultButton2, "In te voegen.  Bent U zeker ?")
            If KtrlBox = vbYes Then
                VrijeTekst = GridText
            End If
    End Select
End If

End Sub

