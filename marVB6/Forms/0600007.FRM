VERSION 5.00
Begin VB.Form LijstIntrastat19 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lijst Intrastat"
   ClientHeight    =   1410
   ClientLeft      =   2880
   ClientTop       =   420
   ClientWidth     =   5895
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
   ScaleHeight     =   1410
   ScaleWidth      =   5895
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGoederen 
      Caption         =   "&Goederenkodes bijwerken"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4560
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   810
      Width           =   1275
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Caption         =   "Afdruk in &Venster"
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   1080
      Width           =   1830
   End
   Begin VB.CommandButton CmdEmailNBB 
      Caption         =   "&E-mail rvListings"
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
      Left            =   4560
      Picture         =   "0600007.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   90
      Width           =   1275
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   390
      Width           =   3240
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   2130
      TabIndex        =   3
      Top             =   690
      Width           =   1200
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "&Afdruk"
      Height          =   675
      Left            =   3600
      Picture         =   "0600007.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   90
      Width           =   900
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
      Height          =   510
      Left            =   3600
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   810
      Width           =   900
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&SubTitel"
      Height          =   270
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   960
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datu&m Drukken"
      Height          =   270
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   690
      Width           =   1395
   End
End
Attribute VB_Name = "LijstIntrastat19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim psTekst(5)              As String
Dim VeldTXT(13)             As String
Dim IntraType               As String * 2

Dim Totaal8                As Double
Dim Totaal9                As Double
Dim Totaal10               As Double

Dim TLijnen                 As Integer
Dim aa                      As String

Private Sub Annuleren_Click()

Unload Me

End Sub

Private Sub CmdEmailNBB_Click()

MsgBox "Enkel beschikbaar voor Servicecontractanten PLUS"

End Sub

Private Sub cmdGoederen_Click()

On Local Error GoTo CancelError
Mim.Teken.filename = PROGRAM_LOCATION + "NOM-N-??.ASC"
Mim.Teken.Filter = "Intrastat ASC-bestand (nom*.asc)|nom*.asc"
Mim.Teken.CancelError = True
Mim.Teken.ShowOpen

Dim Pathnaam As String
Dim PathNaam2 As String
Dim ConnectString As String

Pathnaam = Mim.Teken.filename + " "
Exit Sub

OpenError:
Select Case Err
    Case 3170
        MsgBox Error + vbCrLf + vbCrLf + "Kontroleer type en/of versie van de database a.u.b. !"
    Case Else
        MsgBox Error
End Select

CancelError:

End Sub

Private Sub Drukken_Click()
Dim VorigeSleutel    As String
Dim Lijn            As Integer
Dim XX As Integer

aa = ""
TLijnen = 0
Lijn = 0

    psTekst(2) = "Lijst Intrastat " + IntraType + " " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
    psTekst(0) = TekstLijn(1).text
    psTekst(3) = TekstLijn(0).text

InitVelden
bGetOrGreater TABLE_VARIOUS, 1, vSet(IntraType, 20)
If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) > IntraType Then
    Beep
    SnelHelpPrint "Er zijn geen Intra-communautaire verrichtingen...", BL_LOGGING
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
    bNext TABLE_VARIOUS
    If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) > IntraType Then
        Exit Do
    Else
        GoSub PrintInfo
    End If
Loop
PrintTotaal
Printer.EndDoc
Screen.MousePointer = vbNormal
Annuleren_Click
Exit Sub

PrintInfo:
RecordToVeld TABLE_VARIOUS
If Not DATE_CHECK(vBibTekst(TABLE_VARIOUS, "#v035 #"), PERIODAS_KEY) Then
    Return
ElseIf VorigeSleutel = "" Then
    VorigeSleutel = FVT(TABLE_VARIOUS, 1)
    GoSub EersteZesVelden
ElseIf VorigeSleutel <> KEY_BUF(TABLE_VARIOUS) Then
    PrintTotaal
    VorigeSleutel = FVT(TABLE_VARIOUS, 1)
    GoSub EersteZesVelden
End If

VeldTXT(10) = vBibTekst(TABLE_VARIOUS, "#v033 #")
Select Case Mid(VeldTXT(10), 2, 1)
    Case "0"
        VeldTXT(7) = Dec$(Val(vBibTekst(TABLE_VARIOUS, "#v079 #")), "#######")
        VeldTXT(8) = Dec$(Val(vBibTekst(TABLE_VARIOUS, "#v080 #")), "#######")
        VeldTXT(9) = Dec$(Val(vBibTekst(TABLE_VARIOUS, "#v081 #")), MASK_SY(0))
    Case "1"
        VeldTXT(7) = Dec$(-Val(vBibTekst(TABLE_VARIOUS, "#v079 #")), "#######")
        VeldTXT(8) = Dec$(-Val(vBibTekst(TABLE_VARIOUS, "#v080 #")), "#######")
        VeldTXT(9) = Dec$(-Val(vBibTekst(TABLE_VARIOUS, "#v081 #")), MASK_SY(0))
End Select
Totaal8 = Totaal8 + Val(VeldTXT(7))
Totaal10 = Totaal10 + Val(VeldTXT(9))
PrintVelden
Return

EersteZesVelden:
Lijn = Lijn + 1
VeldTXT(0) = Format(Lijn, "0000")
VeldTXT(1) = vBibTekst(TABLE_VARIOUS, "#v072 #")
VeldTXT(2) = vBibTekst(TABLE_VARIOUS, "#v073 #")
VeldTXT(3) = vBibTekst(TABLE_VARIOUS, "#v074 #")
VeldTXT(4) = vBibTekst(TABLE_VARIOUS, "#v075 #")
VeldTXT(5) = vBibTekst(TABLE_VARIOUS, "#v076 #") + vBibTekst(TABLE_VARIOUS, "#v077 #")
VeldTXT(6) = vBibTekst(TABLE_VARIOUS, "#v078 #")
Return

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Dim T As Integer

Totaal8 = 0
Totaal9 = 0
Totaal10 = 0

IntraType = Format(aIndex)
Caption = Caption + " " + IntraType

TekstLijn(0).text = "Periode : " + BYPERDAT.PeriodeBoekjaar.text
TekstLijn(1).text = MIM_GLOBAL_DATE


End Sub


Private Sub InitVelden()
Dim T As Integer

REPORT_FIELD(0) = "Lijn"
    REPORT_TAB(0) = 1

Select Case IntraType
    Case "19"
        REPORT_FIELD(1) = "Herkomst"
        REPORT_TAB(1) = 6

        REPORT_FIELD(2) = "Oorsprong"
        REPORT_TAB(2) = 16

        REPORT_FIELD(4) = "Plaats lossen"
        REPORT_TAB(4) = 37
    Case "29"
        REPORT_FIELD(1) = "Bestemming"
        REPORT_TAB(1) = 6

        REPORT_FIELD(2) = ""
        REPORT_TAB(2) = 16

        REPORT_FIELD(4) = "Plaats laden"
        REPORT_TAB(4) = 37
End Select

REPORT_FIELD(3) = "Vervoerwijze"
    REPORT_TAB(3) = 26

REPORT_FIELD(5) = "Aard Transactie"
    REPORT_TAB(5) = 48

REPORT_FIELD(6) = "Goederencode"
    REPORT_TAB(6) = 62

REPORT_FIELD(7) = "Netto in KG"
    REPORT_TAB(7) = 75

REPORT_FIELD(8) = "Aanvull. Eenheden"
    REPORT_TAB(8) = 89

REPORT_FIELD(9) = "Stat. waarde BEF"
    REPORT_TAB(9) = 103

REPORT_FIELD(10) = "document"
    REPORT_TAB(10) = 117

REPORT_TAB(11) = 0

If chkAfdrukInVenster Then
    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = Caption
    Xlog.X.Cols = 11
    Xlog.X.Row = 0
    For T = 0 To 10
        Xlog.X.Col = T
        Xlog.X.text = REPORT_FIELD(T)
    Next
    Me.Show
End If

End Sub

Private Sub PrintTitel()
Dim T As Integer

'On Local Error GoTo PrtHandler3

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

Select Case IntraType
    Case "19"
        Printer.Print "Nr.  Lid-Staat Land van  Vervoers-  Plaats van       Aard    Goederencode Nettomassa    Aanvullende   Statistische  dokument"
        Printer.Print "Ln   Herkomst  Oorsprong wijze      lossen     Transactie    (8 CIJFERS)  IN KG.        eenheden      Waarde (BEF)"
    Case "29"
        Printer.Print "Nr.  Lid-Staat           Vervoers-  Plaats van       Aard    Goederencode Nettomassa    Aanvullende   Statistische  dokument"
        Printer.Print "Ln   Bestemming          wijze      laden      Transactie    (8 CIJFERS)  IN KG.        eenheden      Waarde (BEF)"
End Select
Printer.Print FULL_LINE; vbCrLf; vbCrLf;
Exit Sub

PrtHandler3:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub PrintTotaal()
Dim T As Integer

On Local Error GoTo PrtHandler4

For T = 0 To 11
    VeldTXT(T) = ""
Next
VeldTXT(7) = Dec$(Totaal8, "#######")
VeldTXT(9) = Dec$(Totaal10, MASK_SY(0))

TLijnen = TLijnen + 1
If chkAfdrukInVenster Then
    aa = aa + Chr(13)
Else
    Printer.FontBold = True
End If

T = 0
Do While T < 11
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print ; Tab(REPORT_TAB(T));
        Printer.Print ; VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
            aa = aa + Chr(13)
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If chkAfdrukInVenster Then
Else
    Printer.FontBold = False
    Printer.Print vbCrLf;
End If
Totaal8 = 0
Totaal10 = 0
Exit Sub

PrtHandler4:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub PrintVelden()
Dim T As Integer

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
            aa = aa + Chr(13)
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If Printer.CurrentY >= Printer.Height - 1200 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If

End Sub

Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelLength = Len(TekstLijn(Index).text)

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 1
        If DATE_INVALID((TekstLijn(1).text)) Then
            Beep
            TekstLijn(1).text = MIM_GLOBAL_DATE
            TekstLijn(1).SetFocus
        End If
    Case 3
        TekstLijn(3).text = Format(Val(TekstLijn(3).text), "00000")
End Select

End Sub

