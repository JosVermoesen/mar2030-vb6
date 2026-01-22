VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Rappelklanten 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rekeninguittreksels klanten"
   ClientHeight    =   6135
   ClientLeft      =   960
   ClientTop       =   1320
   ClientWidth     =   9285
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
   ScaleHeight     =   6135
   ScaleWidth      =   9285
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Schoonvegen 
      Caption         =   "Schoo&n"
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
      Left            =   7800
      TabIndex        =   15
      Top             =   3600
      Width           =   1440
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2295
      Index           =   0
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   3660
      Width           =   6975
   End
   Begin VB.CheckBox Overschrijvingsstrook 
      Caption         =   "Met OV Strook"
      Height          =   480
      Left            =   7800
      TabIndex        =   12
      Top             =   4800
      Width           =   1065
   End
   Begin VB.CommandButton Samenstellen 
      Caption         =   "Samenstellen"
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
      Left            =   7800
      TabIndex        =   0
      Top             =   120
      Width           =   1445
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1500
      TabIndex        =   1
      Top             =   135
      Width           =   1260
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   3360
      TabIndex        =   2
      Top             =   135
      Width           =   1800
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   5760
      TabIndex        =   3
      Top             =   135
      Width           =   1800
   End
   Begin VB.CheckBox Selektie 
      Caption         =   "Kontrole Vervaldag"
      Enabled         =   0   'False
      Height          =   240
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CheckBox Selektie 
      Caption         =   "Enkel dit boekjaar"
      Enabled         =   0   'False
      Height          =   240
      Index           =   2
      Left            =   2280
      TabIndex        =   5
      Top             =   3240
      Width           =   2055
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Afdrukken"
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
      Left            =   7800
      TabIndex        =   6
      Top             =   3240
      Width           =   1445
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
      Height          =   300
      Left            =   7800
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5640
      Width           =   1445
   End
   Begin MSFlexGridLib.MSFlexGrid grdDokumentDetail 
      Height          =   2595
      Left            =   30
      TabIndex        =   14
      Top             =   540
      Width           =   9225
      _ExtentX        =   16272
      _ExtentY        =   4577
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      SelectionMode   =   1
      AllowUserResizing=   1
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
      Caption         =   "Mededeling"
      Height          =   210
      Index           =   4
      Left            =   6000
      TabIndex        =   8
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Datum kontrole"
      Height          =   240
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Van"
      Height          =   240
      Index           =   2
      Left            =   2880
      TabIndex        =   10
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Tot"
      Height          =   240
      Index           =   3
      Left            =   5280
      TabIndex        =   11
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "Rappelklanten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit


Dim dttot As Double
Dim dtbtw As Double
Dim dvtot As Double

Dim VsoftVanaf          As String
Dim VsoftTot            As String
Dim BeginXbox           As Integer
Dim BeginYbox           As Integer
Dim OVSStrooklijnen As Long

Dim KontaktPersoon As Integer

Private Sub Annuleren_Click()

Unload Me

End Sub

Private Sub RasterSchoon()

With grdDokumentDetail
    .Rows = 2
    .Row = 1
    .Col = 0
    .Text = ""
    
    .Col = 1
    .Text = ""
    
    .Col = 2
    .Text = ""
    
    .Col = 3
    .Text = ""
    
    .Col = 4
    .Text = ""

    .Col = 5
    .Text = ""

    .Col = 6
    .Text = ""

    .Col = 7
    .Text = ""

End With

End Sub





Private Sub DrukAf()

Dim hTAB    As Integer
Dim NogString As String
Dim Nog     As Integer
Dim tSip    As Integer
Dim FlTemp  As Integer
Dim A       As String
Dim dokumentSleutel As String
Dim iRNTxt  As Integer
Dim Teltxt  As Integer
Dim dokumentType As String
ReDim rSip(6) As String
ReDim sSip(6) As String
ReDim VeldInfo(9) As String

ReDim dVeldInfo(7) As Currency
Dim TotaalBedrag As Currency
Dim SubTotaal As Currency

Dim rBTWstr As String
Dim Tekst As String
Dim positieCRLF As Integer
Dim BedragTxt As String
Dim rNTTxt As String
Dim rNTTxt2 As String
Dim ReferteTxt As String


Dim T                   As Integer
Dim TT                  As Integer
Dim Taal                As String
Dim sy                  As String
Dim sy2                 As String

Dim XVan                As Single
Dim XTot                As Single
Dim YVan                As Single
Dim YTot                As Single
Dim X                   As Integer
Dim DummyKlant          As String

If Trim$(grdDokumentDetail.TextMatrix(1, 0)) = "" Then Exit Sub

Set Printer = Printers(dokumentPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "dokumentPRINTER")
If Printer.Orientation = vbPRORLandscape Then
    Printer.Orientation = vbPRORPortrait
    DoEvents
End If

grdDokumentDetail.Row = 1
grdDokumentDetail.Col = 0
If Overschrijvingsstrook.Value Then
    OVSStrooklijnen = 5100
Else
    OVSStrooklijnen = 0
End If

JumpNextKlant:
bGet TABLE_CUSTOMERS, 0, vSet(Mid(grdDokumentDetail.Text, 2), 12)
If Ktrl Then
    MsgBox "KlantFiche bestaat niet meer !"
Else
    RecordToVeld TABLE_CUSTOMERS
    Taal = vBibTekst(TABLE_CUSTOMERS, "#A10C #")
End If

'On Local Error GoTo PrintHandler

    GoSub KopBalk
    GoSub DetailLijnen
    GoSub VoetTekst

grdDokumentDetail.Col = 0
If DummyKlant <> grdDokumentDetail.Text Then
    GoTo JumpNextKlant
Else
    Printer.EndDoc
End If
Exit Sub

KopBalk:
Select Case Taal
    Case "1"
        dokumentType = "Rappèl"
    Case Else
        dokumentType = "Rekeninguittreksel"
End Select
If Val(vBibTekst(TABLE_CUSTOMERS, "#A102 #")) = 0 Then
    rSip(0) = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
Else
    rSip(0) = Mid(fmarBoxText("003", Taal, vBibTekst(TABLE_CUSTOMERS, "#A102 #")), 4, 10) + " " + vBibTekst(TABLE_CUSTOMERS, "#A100 #")
End If
If KontaktPersoon = 1 Then
    If Val(vBibTekst(TABLE_CUSTOMERS, "#vs01 #")) = 0 Then
        rSip(1) = vBibTekst(TABLE_CUSTOMERS, "#A125 #")
    Else
        rSip(1) = Mid(fmarBoxText("003", Taal, vBibTekst(TABLE_CUSTOMERS, "#vs01 #")), 4, 10) + " " + vBibTekst(TABLE_CUSTOMERS, "#A125 #")
    End If
Else
    rSip(1) = ""
End If

rSip(2) = vBibTekst(TABLE_CUSTOMERS, "#A104 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A105 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A106 #")
rSip(4) = vBibTekst(TABLE_CUSTOMERS, "#A109 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A107 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A108 #")
GoSub PrintKopTekst

Printer.FontName = "Courier New"
Printer.FontSize = 10
Printer.FontBold = True

Printer.CurrentY = Val(VsoftVanaf)
Printer.CurrentX = 0

Printer.Print Tab(7);
    XVan = Printer.CurrentX
    YVan = Printer.CurrentY

Printer.Print dokumentType;
    XTot = Printer.CurrentX
    YTot = Printer.CurrentY

Printer.DrawWidth = 4
Printer.Line (XVan - 50, YVan - 50)-(XTot + 50, YTot + 250), , B
    Printer.CurrentX = XVan     'ander geen afdruk met vb4 32bit
    Printer.CurrentY = YVan     'ander geen afdruk met vb4 32bit
    Printer.Print dokumentType; 'ander geen afdruk met vb4 32bit
    Printer.CurrentX = XTot     'ander geen afdruk met vb4 32bit
    Printer.CurrentY = YTot     'ander geen afdruk met vb4 32bit
    Printer.Print vbCrLf;        'ander geen afdruk met vb4 32bit

Printer.Print vbCrLf;
Printer.FontBold = True

Printer.Print Tab(7);
    XVan = Printer.CurrentX
    YVan = Printer.CurrentY

    Printer.Print "KLANT        BTW NUMMER     FAX          DATUM UITTREKSEL";

If vBibTekst(TABLE_CUSTOMERS, "#A161 #") = "" Then
    rBTWstr = ""
Else
    rBTWstr = vBibTekst(TABLE_CUSTOMERS, "#v150 #") + vBibTekst(TABLE_CUSTOMERS, "#A161 #")
End If
Printer.Print vbCrLf; Tab(7); vSet(vBibTekst(TABLE_CUSTOMERS, "#A110 #"), 12); " "; vSet(rBTWstr$, 14); " "; vSet(vBibTekst(TABLE_CUSTOMERS, "#vs02 #"), 12); " "; vSet((TekstLijn(1).Text), 16);
    XTot = Printer.CurrentX
    YTot = Printer.CurrentY

Printer.DrawWidth = 1
Printer.Line (XVan - 50, YVan - 50)-(XTot + 50, YTot + 250), , B
    Printer.CurrentX = XVan     'ander geen afdruk met vb4 32bit
    Printer.CurrentY = YVan     'ander geen afdruk met vb4 32bit
    Printer.Print "KLANT        BTW NUMMER     FAX          DATUM UITTREKSEL";
    Printer.Print vbCrLf; Tab(7); vSet(vBibTekst(TABLE_CUSTOMERS, "#A110 #"), 12); " "; vSet(rBTWstr$, 14); " "; vSet(vBibTekst(TABLE_CUSTOMERS, "#vs02 #"), 12); " "; vSet((TekstLijn(1).Text), 16);
    Printer.CurrentX = XTot     'ander geen afdruk met vb4 32bit
    Printer.CurrentY = YTot     'ander geen afdruk met vb4 32bit
    Printer.Print vbCrLf;        'ander geen afdruk met vb4 32bit

Printer.FontBold = False
Printer.Print vbCrLf;

Printer.Print Tab(7);
    XVan = Printer.CurrentX
    YVan = Printer.CurrentY
    Printer.Print "dokument          TOTAAL DATUM      VERVALDAG       BETAALD         REST FIN.STUK";
    XTot = Printer.CurrentX
    YTot = Printer.CurrentY

Printer.DrawWidth = 1
Printer.Line (XVan - 50, YVan - 50)-(XTot + 50, YTot + 250), , B
    Printer.CurrentX = XVan     'ander geen afdruk met vb4 32bit
    Printer.CurrentY = YVan     'ander geen afdruk met vb4 32bit
    Printer.Print "dokument          TOTAAL DATUM      VERVALDAG       BETAALD         REST FIN.STUK";
    Printer.CurrentX = XTot     'ander geen afdruk met vb4 32bit
    Printer.CurrentY = YTot     'ander geen afdruk met vb4 32bit
    Printer.Print vbCrLf;        'ander geen afdruk met vb4 32bit

Printer.Print vbCrLf;
Printer.FontBold = True
Return

DetailLijnen:
grdDokumentDetail.Col = 0
DummyKlant = grdDokumentDetail.Text
SubTotaal = 0
Do
    grdDokumentDetail.Col = 7
    Printer.Print Tab(7); grdDokumentDetail.Text; " ";
    grdDokumentDetail.Col = 1
    Printer.Print Dec(CDbl(grdDokumentDetail.Text), MASK_EURBH); " ";
    grdDokumentDetail.Col = 2
    Printer.Print grdDokumentDetail.Text; " ";
    grdDokumentDetail.Col = 3
    Printer.Print grdDokumentDetail.Text; " ";
    grdDokumentDetail.Col = 4
    Printer.Print Dec(CDbl(grdDokumentDetail.Text), MASK_EURBH); " ";
    grdDokumentDetail.Col = 5
    Printer.Print Dec(CDbl(grdDokumentDetail.Text), MASK_EURBH); " ";
    SubTotaal = SubTotaal + CDbl(grdDokumentDetail.Text)
    grdDokumentDetail.Col = 6
    Printer.Print grdDokumentDetail.Text; " "; vbCrLf;
    If Printer.CurrentY >= Val(VsoftTot) - 1550 - OVSStrooklijnen Then
        GoSub OnderKant
        GoSub VoetTekst
        GoSub KopBalk
    End If
    If grdDokumentDetail.Row = grdDokumentDetail.Rows - 1 Then
        Exit Do
    Else
        grdDokumentDetail.Row = grdDokumentDetail.Row + 1
    End If
    grdDokumentDetail.Col = 0
    If DummyKlant <> grdDokumentDetail.Text Then
        Exit Do
    End If
Loop

OnderKant:
Printer.Print vbCrLf; Tab(7); "TOTAAL IN " + Mim.SnelHelp.Panels(2) + " : "; Tab(67); Dec$((SubTotaal), MASK_EURBH); vbCrLf; vbCrLf;
'Hier nog de tekst ook

Tekst = RTrim$(TekstLijn(0).Text)
Printer.FontName = TekstLijn(0).FontName
Do While Tekst <> ""
    Printer.CurrentX = XVan
    positieCRLF = InStr(Tekst, vbCrLf)
    If positieCRLF Then
        Printer.Print Left(Tekst, positieCRLF - 1); vbCrLf;
        Tekst = Right(Tekst, Len(Tekst) - positieCRLF - 1)
    Else
        Printer.Print Tekst; vbCrLf;
        Tekst = ""
    End If
    If Printer.CurrentY >= Val(VsoftTot) - 100 Then
        MsgBox "Meer berichtgeving dan mogelijk op één blad"
    End If
Loop
Return

PrintKopTekst:
PrintUserDef "1" + Taal + "5"
Printer.CurrentX = BeginXbox + 350
Printer.CurrentY = BeginYbox + 350
Printer.FontSize = 10
Printer.FontBold = True
For tSip = 0 To 4
    Printer.Print UCase$(rSip(tSip)); vbCrLf;
    Printer.CurrentX = BeginXbox + 350
Next
Return

VoetTekst:
If OVSStrooklijnen Then
    GoSub Overschrijvingsstrook
End If
Printer.NewPage
Return

Overschrijvingsstrook:
If Dir$(LOCATION_COMPANYDATA + "LicMarnt.###") = "" Then
    Beep
    MsgBox LOCATION_COMPANYDATA + "LicMarnt.### niet te vinden !  Hierna wordt kladblok opgestart.  Breng uw eigen gegevens in a.u.b. !", 0, "Foutieve Installatie ?"

    On Local Error Resume Next
    X = Shell("notepad.exe " + LOCATION_COMPANYDATA + "LicMarnt.###", 1)
    Return
Else
    FlTemp = FreeFile
    Open LOCATION_COMPANYDATA + "LicMarnt.###" For Input As FlTemp
        Line Input #FlTemp, sSip(0)
        Line Input #FlTemp, sSip(1)
        Line Input #FlTemp, sSip(2)
        Line Input #FlTemp, sSip(3)
        Line Input #FlTemp, sSip(4)
        Line Input #FlTemp, sSip(5)
        Line Input #FlTemp, A$
        Err = 0
        On Local Error Resume Next
        Line Input #FlTemp, NogString
        If Err Then
            Nog = 0
        Else
            Nog = Val(NogString)
        End If
    Close FlTemp
    hTAB = Val(A$)
End If
On Local Error GoTo 0

Printer.FontName = "Courier New"
Printer.FontSize = 12
Printer.FontBold = True
Printer.CurrentY = 11000
Printer.CurrentX = 0
For T = 1 To 1 + Nog
    Printer.Print vbCrLf;
Next

If bhEuro Then
    BedragTxt = "*" + Format(SubTotaal, MASK_EUR) + "*"
Else
    BedragTxt = "*" + Format(SubTotaal) + "*"
End If

NieuwFormulier:
If bhEuro Then
    rNTTxt = Dec((SubTotaal), "#######0.00")
    GoSub Spatieren
    Mid(rNTTxt2, 17, 1) = " "
Else
    rNTTxt = Dec((SubTotaal), MASK_SY(0))
    GoSub Spatieren
End If
Printer.Print Tab(10 + hTAB); BedragTxt;
Printer.Print Tab(52 + hTAB); rNTTxt2; vbCrLf; vbCrLf;

rNTTxt = Mid(sSip(0), 1, 3) + Mid(sSip(0), 5, 7) + Mid(sSip(0), 13, 2)
GoSub Spatieren
rNTTxt2 = Mid(rNTTxt2, 1, 6) + " " + Mid(rNTTxt2, 7, 14) + " " + Mid(rNTTxt2, 21, 3)
Printer.Print Tab(52 + hTAB); rNTTxt2;
Printer.FontBold = False
Printer.Print vbCrLf; vbCrLf;

For T = 0 To 4
    Printer.Print Tab(25 + hTAB); Left(UCase$(rSip(T)), 26); Tab(52 + hTAB); UCase$(sSip(T + 1)); vbCrLf;
Next

rNTTxt = dokumentSleutel
    GoSub Spatieren
    Printer.FontBold = True
    Printer.Print vbCrLf; Tab(1 + hTAB); rNTTxt2; Tab(52 + hTAB); rNTTxt2;
    Printer.FontBold = False
Return

Spatieren:
iRNTxt = Len(rNTTxt)
rNTTxt2 = ""
For Teltxt = 1 To iRNTxt
    rNTTxt2 = rNTTxt2 + Mid(rNTTxt, Teltxt, 1) + " "
Next
Return

PrintHandler:
MsgBox "Stopkode extern toestel." + vbCrLf + vbCrLf + "Kontroleer, of rapporteer aan Vsoft.", 0, "Onverwachte situatie"
Resume

FileHandler:
MsgBox "Onverwachte fout bij openen bestand." + vbCrLf + vbCrLf + "Stopkode : " + Format(Err) + vbCrLf + vbCrLf + "Rapporteer aan Vsoft of kontroleer uw setup."
Ktrl = 100
Exit Sub
Resume

End Sub

Private Sub Drukken_Click()

MousePointer = 11
DrukAf
MousePointer = 0
Annuleren.SetFocus

End Sub

Private Sub Form_Load()
Dim TempB_MODUS As Integer

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

If BJPERDAT.Boekjaar.ListIndex <> 0 Then
    MsgBox "Eerst hoogste boekjaar activeren, a.u.b.", vbExclamation
    Unload Me
    Exit Sub
End If
'enkel leesmodus openen
dttot = 0
dtbtw = 0
dvtot = 0
TekstLijn(1).Text = MIM_GLOBAL_DATE

TekstLijn(2).Text = "0"
TekstLijn(3).Text = String$(12, "z")
grdDokumentDetail.Rows = 2
grdDokumentDetail.Cols = 8
grdDokumentDetail.Col = 0
grdDokumentDetail.Row = 0
grdDokumentDetail.Text = "Klantkode"
grdDokumentDetail.Col = 1
grdDokumentDetail.Text = "Totaal"
grdDokumentDetail.Col = 2
grdDokumentDetail.Text = "Datum"
grdDokumentDetail.Col = 3
grdDokumentDetail.Text = "Vervaldag"
grdDokumentDetail.Col = 4
grdDokumentDetail.Text = "Betaald"
grdDokumentDetail.Col = 5
grdDokumentDetail.Text = "Rest"
grdDokumentDetail.Col = 6
grdDokumentDetail.Text = "Fin.Stuk"
grdDokumentDetail.Col = 7
grdDokumentDetail.Text = "dokument"

grdDokumentDetail.ColWidth(0) = 1455
grdDokumentDetail.ColWidth(1) = 1130
grdDokumentDetail.ColWidth(2) = 975
grdDokumentDetail.ColWidth(3) = 975
grdDokumentDetail.ColWidth(4) = 1130
grdDokumentDetail.ColWidth(5) = 1130
grdDokumentDetail.ColWidth(6) = 885
grdDokumentDetail.ColWidth(7) = 1185
KontaktPersoon = Val(String99(READING, 201))
If Dir$(LOCATION_COMPANYDATA + "Ruit.txt") <> "" Then
    Dim FlTemp As Integer
    Dim aa As String
    Dim AAA As String
    FlTemp = FreeFile
    Open LOCATION_COMPANYDATA + "RUIT.TXT" For Input As FlTemp
    Do While Not EOF(FlTemp)
        Line Input #FlTemp, aa
        AAA = AAA + aa + vbCrLf
    Loop
    Close FlTemp
    TekstLijn(0).Text = AAA
End If

End Sub


Private Sub GrddokumentDetail_Click()

GrddokumentDetail_KeyUp 40, 0

End Sub

Private Sub GrddokumentDetail_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 46
        If grdDokumentDetail.Rows > 2 Then
            grdDokumentDetail.RemoveItem grdDokumentDetail.Row
            GrddokumentDetail_Click
        End If
End Select

End Sub

Private Sub grdDokumentDetail_KeyPress(KeyAscii As Integer)

    'GrddokumentDetail_KeyUp 40, 0

End Sub

Private Sub GrddokumentDetail_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 38, 40
        grdDokumentDetail.Col = 0
        bGet TABLE_CUSTOMERS, 0, vSet(Mid(grdDokumentDetail.Text, 2), 12)
        If Ktrl Then
            'MsgBox "KlantFiche bestaat niet meer !"
            SnelHelpPrint "Ogenblik..", BL_LOGGING
        Else
            RecordToVeld TABLE_CUSTOMERS
            SnelHelpPrint vBibTekst(TABLE_CUSTOMERS, "#A100 #") + ", " + vBibTekst(TABLE_CUSTOMERS, "#A104 #") + ", " + vBibTekst(TABLE_CUSTOMERS, "#A108 #"), BL_LOGGING
        End If
End Select

End Sub

Private Sub KTRLBalans(Fl As Integer)
Dim Cumul       As Currency
Dim dTotaal     As Currency
Dim dBetaald    As Currency

Dim aa          As String
Dim AAA         As String
Dim T           As Integer
Dim VoorLetter  As String * 1

aa = ""
T = 0
bGetOrGreater TABLE_INVOICES, 1, vSet("K" + TekstLijn(2).Text, 13)
If Ktrl Or KEY_BUF(TABLE_INVOICES) > vSet("K" + TekstLijn(3).Text, 13) Then
    Beep
    MsgBox "Selektie buiten mogelijke dokumenten"
    Exit Sub
Else
    RecordToVeld TABLE_INVOICES
End If
If vSet("K" + TekstLijn(2).Text, 13) = vSet("K" + TekstLijn(3).Text, 13) Then
    If vSet(KEY_BUF(TABLE_INVOICES), 13) <> vSet("K" + TekstLijn(2).Text, 13) Then
        Beep
        MsgBox "Geen dokumenten voor " + vBibTekst(Fl, "#A100 #")
        Exit Sub
    End If
End If
Screen.MousePointer = vbHourglass
grdDokumentDetail.Rows = 1
grdDokumentDetail.Refresh
GoSub VolgendeLijn
Do
    bNext TABLE_INVOICES
    If Ktrl Or KEY_BUF(TABLE_INVOICES) > vSet("K" + TekstLijn(3).Text, 13) Then
        Exit Do
    Else
        RecordToVeld TABLE_INVOICES
        GoSub VolgendeLijn
    End If
Loop
Screen.MousePointer = vbNormal
On Error Resume Next
grdDokumentDetail.Row = 1
grdDokumentDetail.Col = 0
Exit Sub

VolgendeLijn:
dBetaald = Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
If Mid(vBibTekst(TABLE_INVOICES, "#v033 #"), 2, 1) = "1" Then
    dTotaal = -dTotaal
    dBetaald = -dBetaald
End If

If dBetaald = dTotaal Then
Else
    AAA = vBibTekst(TABLE_INVOICES, "#v034 #") & vbTab
    AAA = AAA & Format(dTotaal, "#,##0.00") & vbTab
    AAA = AAA & DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v035 #")) & vbTab
    AAA = AAA & DATE_TEXT(vBibTekst(TABLE_INVOICES, "#v036 #")) & vbTab
    AAA = AAA & Format(dBetaald, "#,##0.00") & vbTab
    AAA = AAA & Format(dTotaal - dBetaald, "#,##0.00") & vbTab
    AAA = AAA & vSet(vBibTekst(TABLE_INVOICES, "#v038 #"), 8) & vbTab
    AAA = AAA & vBibTekst(TABLE_INVOICES, "#v033 #")
    T = T + 1
    grdDokumentDetail.AddItem AAA
    grdDokumentDetail.Refresh
End If
Return

End Sub

Private Sub PrintUserDef(TypeEnTaal As String)
Dim T               As Integer
Dim Tabulatie       As Integer
Dim FlFree          As Integer
Dim Teller          As Integer
Dim Dummy           As String

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

If Dir$(LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + ".Txt") = "" Then
    Beep
    MsgBox "Dokumenlay-out : " + LOCATION_COMPANYDATA + "Dok" + TypeEnTaal + ".Txt niet te vinden !"
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

Private Sub grdDokumentDetail_RowColChange()

    'GrddokumentDetail_KeyUp 40, 0

End Sub

Private Sub Samenstellen_Click()

If TekstLijn(2) <> TekstLijn(3) Then
    Msg = "Samenstellen van " & TekstLijn(2) & " tot " & TekstLijn(3)
    Msg = Msg & vbCr & vbCr & "Bent U zeker"
    KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
    If KtrlBox = vbNo Then
        Exit Sub
    End If
End If
KTRLBalans TABLE_CUSTOMERS

End Sub

Private Sub SchoonVegen_Click()

 RasterSchoon
 
End Sub

Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelLength = Len(TekstLijn(Index).Text)
Select Case Index
    Case 0
        Samenstellen.Default = False
    Case 2, 3
        SnelHelpPrint "Dubbelklikken of [Ctrl] voor geïndexeerd zoeken", BL_LOGGING
        Samenstellen.Default = True
End Select

End Sub

Private Sub TekstLijn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case Index
    Case 2, 3
        Select Case KeyCode
            Case 17
                aIndex = 1
                SharedFl = TABLE_CUSTOMERS
                GridText = ""
                SqlSearch.Show 1
                If Ktrl = 0 Then
                    TekstLijn(Index).Text = vBibTekst(TABLE_CUSTOMERS, "#A110 #")
                    If Index = 2 Then
                        TekstLijn(3).Text = TekstLijn(2).Text
                    End If
                End If
        End Select
End Select

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)
Dim FlTemp As Integer

Select Case Index
    Case 0
        Samenstellen.Default = True
        FlTemp = FreeFile
        Open LOCATION_COMPANYDATA + "RUIT.TXT" For Output As FlTemp
            Print #FlTemp, TekstLijn(0).Text;
        Close FlTemp
    Case 1
        Samenstellen.Default = False
        If DATE_INVALID((TekstLijn(1).Text)) Then
            Beep
            TekstLijn(1).Text = MIM_GLOBAL_DATE
            TekstLijn(1).SetFocus
        End If
    Case 2
        Samenstellen.Default = False
End Select

End Sub

