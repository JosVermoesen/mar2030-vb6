VERSION 5.00
Begin VB.Form ProduktieBAPrive 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produktie B.A. Prive Leven"
   ClientHeight    =   2340
   ClientLeft      =   960
   ClientTop       =   1590
   ClientWidth     =   8145
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0C0&
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2340
   ScaleWidth      =   8145
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1440
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   1260
      Width           =   2160
   End
   Begin VB.CommandButton Versturen 
      Caption         =   "Versturen"
      Height          =   420
      Left            =   6720
      TabIndex        =   6
      Top             =   480
      Width           =   1320
   End
   Begin VB.CommandButton Schoonvegen 
      Caption         =   "Schoon"
      Height          =   420
      Left            =   6720
      TabIndex        =   7
      Top             =   900
      Width           =   1320
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   420
      Left            =   6720
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1320
   End
   Begin VB.TextBox Referte 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   1440
   End
   Begin VB.ComboBox MaatschappijKeuze 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin VB.ComboBox VerzoekType 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1800
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.CommandButton Paginas 
      Caption         =   "Pagina"
      Enabled         =   0   'False
      Height          =   420
      Index           =   0
      Left            =   6720
      TabIndex        =   5
      Top             =   60
      Width           =   1320
   End
   Begin VB.CheckBox Selektie 
      Caption         =   "Enkel AssurNet toepasbare maatschappijen"
      Height          =   240
      Left            =   0
      TabIndex        =   17
      Top             =   1920
      Width           =   4035
   End
   Begin VB.Label Label 
      Caption         =   "Maatschappij"
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   300
      Width           =   1440
   End
   Begin VB.Label Label 
      Caption         =   "Klant"
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   10
      Top             =   900
      Width           =   1440
   End
   Begin VB.Label Label 
      Caption         =   "Type verzoek"
      Height          =   240
      Index           =   5
      Left            =   120
      TabIndex        =   12
      Top             =   600
      Width           =   1440
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      Caption         =   "Status"
      Height          =   240
      Index           =   3
      Left            =   5700
      TabIndex        =   15
      Top             =   1260
      Width           =   855
   End
   Begin VB.Label Label 
      Caption         =   "Referte"
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1500
      Width           =   1440
   End
   Begin VB.Label Label 
      Caption         =   "PolisNummer"
      Height          =   300
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Label Status 
      Height          =   240
      Index           =   0
      Left            =   5940
      TabIndex        =   11
      Top             =   1500
      Width           =   600
   End
   Begin VB.Label KlantNaam 
      Enabled         =   0   'False
      Height          =   240
      Left            =   3300
      TabIndex        =   16
      Top             =   960
      Width           =   3255
   End
End
Attribute VB_Name = "ProduktieBAPrive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim Pagina1 As String * 740

Dim DataPos                As Integer
Dim DataLen                As Integer
Dim OptieTxt               As String
Dim TempoTxt               As String
Dim CrText                 As String
Dim CrText2                As String

Dim PolisNummer            As String * 12
Dim Maatschappij           As String * 4

Private Sub Annuleren_Click()

Unload Me

End Sub


Private Sub Form_Load()

PaginaTeller = 0
VerzoekType.AddItem "067: Nieuwe zaak B.A. Prive Leven"
VerzoekType.ListIndex = 0

InstallMij 0
GridText = ""
Schoon

End Sub


Private Sub InstallMij(MijVlag As Integer)
Dim T As Integer

MaatschappijKeuze.Clear

T = -1
bGetOrGreater FlLeverancier, 0, "CO"
If Ktrl Then
    Beep
    Exit Sub
Else
    RecordToVeld FlLeverancier
End If
Screen.MousePointer = vbHourglass
If Left(KeyBuf(FlLeverancier), 2) <> "CO" Then
    Beep
Else
    If MijVlag Then
        If KTRLVerzoek(Left(VerzoekType.Text, 3), Right(KeyBuf(FlLeverancier), 4)) Then
            T = T + 1
            MaatschappijKeuze.AddItem Mid(vBibTekst(FlLeverancier, "#A110 #"), 3, 4) + ": " + vBibTekst(FlLeverancier, "#A100 #") + "/" + vSet(vBibTekst(FlLeverancier, "#A400 #"), 8)
        End If
    Else
        T = T + 1
        MaatschappijKeuze.AddItem Mid(vBibTekst(FlLeverancier, "#A110 #"), 3, 4) + ": " + vBibTekst(FlLeverancier, "#A100 #") + "/" + vSet(vBibTekst(FlLeverancier, "#A400 #"), 8)
    End If

    Do
        bNext FlLeverancier
        If Ktrl Or Left(KeyBuf(FlLeverancier), 2) <> "CO" Then
            Exit Do
        Else
            RecordToVeld FlLeverancier
            If MijVlag Then
                If KTRLVerzoek(Left(VerzoekType.Text, 3), Right(KeyBuf(FlLeverancier), 4)) Then
                    T = T + 1
                    MaatschappijKeuze.AddItem Mid(vBibTekst(FlLeverancier, "#A110 #"), 3, 4) + ": " + vBibTekst(FlLeverancier, "#A100 #") + "/" + vSet(vBibTekst(FlLeverancier, "#A400 #"), 8)
                End If
            Else
                T = T + 1
                MaatschappijKeuze.AddItem Mid(vBibTekst(FlLeverancier, "#A110 #"), 3, 4) + ": " + vBibTekst(FlLeverancier, "#A100 #") + "/" + vSet(vBibTekst(FlLeverancier, "#A400 #"), 8)
            End If
        End If
    Loop
End If
MaatschappijKeuze.ListIndex = 0
Screen.MousePointer = vbNormal

End Sub

Private Sub Paginas_Click(Index As Integer)
Dim fSleutel        As String
Dim MijNaam         As String
Dim TeleBibTekst    As String
Dim Msg             As String
Dim fKtrl           As Integer
Dim T               As Integer
Dim fIndex          As Integer

Load Xlog
Xlog.Left = 0
Xlog.Width = Screen.Width
Xlog.X.Width = Screen.Width - 2
Xlog.Height = Xlog.Height + 2
'Xlog.InfoLabel.Caption = "vsfTeleBib Omschrijving                   Variabelen"

TeleBibPagina "PBAPRIVE.DEF"
Pagina1 = ""

Maatschappij = Left(MaatschappijKeuze.Text, 4)
PolisNummer = TekstInfo(1).Text

If TekstInfo(1).Text = "" Then
    fIndex = 0
    fSleutel = TekstInfo(0).Text
Else
    fIndex = 1
    fSleutel = "4" + Maatschappij + PolisNummer
End If

TLBRecord(FlAllerlei) = ""
GridText = ""
bGet FlAllerlei, fIndex, fSleutel
If Ktrl Then
    MsgBox "Geen gestruktureerde informatie te vinden", 0, TekstInfo(0).Text + " / " + PolisNummer
    fKtrl = 99
    bGet FlLeverancier, 0, "CO" + Maatschappij
    If Ktrl Then
    Else
        RecordToVeld FlLeverancier
    End If
    'vBib Fl, "string", "v023"
    vBib FlAllerlei, PolisNummer, "A000"
    'vBib FlAllerlei, FVT(FlLeverancier, 11), "A400"
    'vBib FlAllerlei  FVT(FlKlant, 1), "A100"
    'vBib FlAllerlei  (Referte.Text), "A110"
    If PolisNummer = Space$(6) Then
    Else
        'vBib FlAllerlei Mid(FVT(FlPolis, 7), 7, 2) + Mid(FVT(FlPolis, 7), 5, 2), "A320"
    End If
Else
    fKtrl = 100
    RecordToVeld FlAllerlei
    bDelete FlAllerlei
    Do
        bNext FlAllerlei
        If KeyBuf(FlAllerlei) = FVT(FlAllerlei, fIndex) Then
            RecordToVeld FlAllerlei
            bDelete FlAllerlei
        Else
            Exit Do
        End If
    Loop
End If

T = 0
Do While TeleBibCode(T) <> Space$(10)
    'Stop
    'CrText = TeleBibTxt("#" + Mid(TeleBibCode(T), 5, 5) + "#")
    'Stop
    'PositieX = Xlog.X.ListCount
    'Xlog.X.AddItem TeleBibCode(T) + Chr$(179) + TeleBibTekst(T) + Chr$(179) + CrText, PositieX
    'T = T + 1
Loop

XLogShow:
Xlog.Caption = VerzoekType.Text
'Stop
'Xlog.X.ListIndex = 0
Xlog.WijzigenLijn.Default = True
Xlog.Afsluiten.TabStop = True
XLogKey = ""

ShowIt

End Sub

Private Sub Schoon()
Dim T As Integer

For T = 0 To 1
    TekstInfo(T).Text = ""
    TekstInfo(T).Enabled = True
Next

Paginas(0).Enabled = False
Status(0).Caption = ""

KlantNaam.Caption = ""
Referte.Text = ""
MaatschappijKeuze.Enabled = True
XLogKey = ""

End Sub

Private Sub SchoonVegen_Click()

Unload Xlog
Schoon
MaatschappijKeuze.SetFocus

End Sub

Private Sub SchrijfWeg()
Dim FlNetTemp As Integer
On Local Error GoTo SchrijfWegError

FlNetTemp = FreeFile
Open AssurnetLokatie + "AS1.SND" For Append As FlNetTemp

Print #FlNetTemp, Pagina1; vbCrLf;

Close FlNetTemp
    Exit Sub

SchrijfWegError:
    MsgBox "Openen van " + AssurnetLokatie + "AS1.SND onmogelijk.  Probeer opnieuw"
    Resume

End Sub

Private Sub Selektie_Click()

InstallMij (Selektie.Value)

End Sub

Private Sub ShowIt()

'Stop
Xlog.Show
ProduktieBAPrive.Hide

End Sub

Private Sub TekstInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 17
        Select Case Index
            Case 0
                Ktrl = 99
                'Stop
                'Unload Klantenfiche
                'CLIPBOARD.SetText TekstInfo(Index).Text
                BasisB(5).Show 1

                'bGet FlKlant, 0, "010018"
                If Ktrl Then
                    Beep
                Else
                    RecordToVeld FlKlant
                    TekstInfo(Index).Text = FVT(FlKlant, 0)
                    Referte.Text = FVT(FlKlant, 0)
                    KlantNaam.Caption = FVT(FlKlant, 1)
                    TekstInfo(1).Text = ""
                    Paginas(0).Enabled = True
                    ProduktieBAPrive.SetFocus
                    Paginas(0).SetFocus
                End If
                ProduktieBAPrive.SetFocus

            Case 1
                Stop
                'Unload PolisFiche
                Ktrl = 99
                Clipboard.SetText TekstInfo(Index).Text
                'PolisFiche.Show 1
                TekstInfo(Index).Text = FVT(FlPolis, 3)
                ProduktieBAPrive.SetFocus

                'bGet FlPolis, 3, "03915843    "
                If Ktrl Then
                    Beep
                    TekstInfo(1).Text = ""
                    TekstInfo(0).Enabled = True
                Else
                    RecordToVeld FlPolis
                    If Left(FVT(FlPolis, 9), 1) <> "4" Then
                        MsgBox "Geen kontrakt Prive Leven !" + vbCrLf + vbCrLf + "Kontroleer waarborgkode 1"
                        TekstInfo(1).Text = ""
                        TekstInfo(0).Enabled = True
                    Else
                        T = MaatschappijKeuze.ListCount
                        For TT = 1 To T
                            MaatschappijKeuze.ListIndex = TT - 1
                            If FVT(FlPolis, 2) = Left(MaatschappijKeuze.Text, 4) Then
                                MaatschappijKeuze.Enabled = False
                                Exit For
                            End If
                        Next

                        bGet FlKlant, 0, FVT(FlPolis, 1)
                        If Ktrl Then
                        Else
                            TekstInfo(1).Text = FVT(FlPolis, 3)
                            TekstInfo(0).Enabled = False
                            RecordToVeld FlKlant
                            TekstInfo(0).Text = FVT(FlKlant, 0)
                            KlantNaam.Caption = FVT(FlKlant, 1)
                            Referte.Text = FVT(FlKlant, 0)
                            Paginas(0).Enabled = True
                            Paginas(0).SetFocus
                        End If
                    End If
                End If
        End Select
End Select

End Sub

Private Sub TeleBibPagina(fDefinitie As String)
Dim FlTemp As Integer
Dim T As Integer
Dim TempoVlag As String
Dim bList As Integer

'code= 1 - 1 : Poliskontrole 1 = ON
'      2 - 3 : Selektiekeuze waarde (ListIDX)
'      4 - 4 :
'      5 - 8 : TeleBib
'      9 - 9 : Volgnummer eventueel 1 tot 9
'      10    : * voor verplichte invulling

bList = VerzoekType.ListIndex + 1
FlTemp = FreeFile
Open ProgrammaLokatie + "Def\" + fDefinitie For Input As FlTemp
T = 0
Do While Not EOF(FlTemp)
    Input #FlTemp, TeleBibCode(T), TeleBibTekst(T), TeleBibType(T), TeleBibLengte(T), TeleBibPos(T)
    T = T + 1
Loop
Close FlTemp
TeleBibCode(T) = ""

End Sub

Private Sub Versturen_Click()
Dim PolisNummer     As String * 12
Dim MijNaam         As String
Dim TypeBericht     As String
Dim TeleBibTekst    As String
Dim Msg             As String
Dim OptieTxt        As String
Dim TempoTxt        As String
Dim CrText          As String
Dim CrText2         As String
Dim fKtrl           As Integer
Dim T               As Integer
Dim DataPos         As Integer
Dim DataLen         As Integer

End Sub

Private Sub VoorVoegsel()

Pagina1 = ""
Mid(Pagina1, 1, 3) = "AS1"                               'constante
Mid(Pagina1, 16, 1) = "0"                                'constante
Mid(Pagina1, 17, 3) = Left(VerzoekType.Text, 3)        'Klient of auto
Mid(Pagina1, 20, 1) = "1"                                'gestruktureerde inhoud
Mid(Pagina1, 21, 1) = "1"                                'producent
Mid(Pagina1, 22, 8) = vSet(ProducentNummer, 8)        'ASNET prod.nr.
Mid(Pagina1, 30, 12) = Format(Now, "yymmddhhmmss")      'datum
Mid(Pagina1, 42, 1) = "1"                                'gewoon niet vertrouwelijk
Mid(Pagina1, 43, 1) = "2"                                'bestemmeling=mij.
Mid(Pagina1, 44, 4) = Left(MaatschappijKeuze.Text, 4)        'BVVO kode
Mid(Pagina1, 48, 2) = "01"                               'volgnummer
Mid(Pagina1, 50, 1) = "X"                                'vervolg tekst
Mid(Pagina1, 104, 1) = "A"                               'dienstkode
Mid(Pagina1, 105, 1) = "4"                               'takkode

End Sub

