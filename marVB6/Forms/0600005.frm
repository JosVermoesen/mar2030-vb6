VERSION 5.00
Begin VB.Form DiversePostenBoek 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Diverse postenboek"
   ClientHeight    =   930
   ClientLeft      =   525
   ClientTop       =   2370
   ClientWidth     =   6195
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
   ScaleHeight     =   930
   ScaleWidth      =   6195
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkAfdrukLiggend 
      Alignment       =   1  'Right Justify
      Caption         =   "Liggende Printerafdruk"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   630
      Width           =   2310
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   630
      Value           =   1  'Checked
      Width           =   1830
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   4320
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   2850
      TabIndex        =   1
      Top             =   0
      Width           =   1560
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&drukken"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   30
      Width           =   1440
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "&Sluiten"
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
      Left            =   4680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   450
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sub&Titel"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   60
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Datu&m Drukken"
      Height          =   195
      Index           =   1
      Left            =   1440
      TabIndex        =   5
      Top             =   30
      Width           =   1335
   End
End
Attribute VB_Name = "DiversePostenBoek"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim psTekst(5)  As String
Dim TotaalD     As Double
Dim TotaalC     As Double

Dim VeldTXT(17) As String

Dim TLijnen     As Integer
Dim aa          As String

Dim PlGrensVan  As String * 8
Dim PlGrensTot  As String * 8

Private Sub Annuleren_Click()
            
Unload Me

End Sub

Private Sub CumulPrint()
Dim RekeningNaam As String * 30
Dim Tabul As Integer

If chkAfdrukInVenster Then
Else
    On Local Error Resume Next
    Printer.EndDoc
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
    Err = 0
    On Local Error Resume Next
    If Err - 480 > 0 Then Exit Sub
    PrintTitel
    PrintTotaal
End If

If chkAfdrukInVenster Then
Else
    Printer.Print
    Printer.Print Tab(2); "** CENTRALISATIE/VIERKANTSCONTROLE **"
    Printer.Print
End If

bClose TABLE_DUMMY
bFirst TABLE_DUMMY, 0
RecordToVeld TABLE_DUMMY
bGet TABLE_LEDGERACCOUNTS, 0, Left(FVT(TABLE_DUMMY, 0), 7)
If Ktrl Then
    RekeningNaam = "Rekening reeds vernietigd !!!"
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    RekeningNaam = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
End If
Tabul = 0
If chkAfdrukInVenster Then
Else
    Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")), "####"); " x "; vSet(vBibTekst(TABLE_DUMMY, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")), MASK_EURBH);
Do
    bNext TABLE_DUMMY
    If Ktrl Then
        Exit Do
    End If
    RecordToVeld TABLE_DUMMY
    bGet TABLE_LEDGERACCOUNTS, 0, Left(FVT(TABLE_DUMMY, 0), 7)
    If Ktrl Then
        RekeningNaam = "Rekening reeds vernietigd !!!"
    Else
        RecordToVeld TABLE_LEDGERACCOUNTS
        RekeningNaam = vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #")
    End If
    If chkAfdrukInVenster Then
    ElseIf Tabul = 0 Then
        Tabul = 58
        Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")), "####"); " x "; vSet(vBibTekst(TABLE_DUMMY, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")), MASK_EURBH); vbCrLf;
    Else
        Tabul = 0
        Printer.Print Tab(Tabul + 2); Dec$(Val(vBibTekst(TABLE_DUMMY, "#v013 #")), "####"); " x "; vSet(vBibTekst(TABLE_DUMMY, "#v089 #"), 7); " "; RekeningNaam; " "; Dec$(Val(vBibTekst(TABLE_DUMMY, "#v068 #")), MASK_EURBH);
    End If
Loop
End If

If chkAfdrukInVenster Then
    Xlog.X.Row = 1
    Xlog.X.Col = 0
    Xlog.X.ColWidth(0) = 555
    Xlog.X.ColWidth(1) = 1050
    Xlog.X.ColWidth(2) = 780
    Xlog.X.ColWidth(3) = 3150
    Xlog.X.ColWidth(4) = 2370
    Xlog.X.ColWidth(5) = 1125
    Xlog.X.ColWidth(6) = 1125
    Xlog.X.ColWidth(7) = 870
    Xlog.WindowState = 0
    
    Xlog.X.ColAlignment(5) = flexAlignRightTop
    Xlog.X.ColAlignment(6) = flexAlignRightTop
    Screen.MousePointer = vbNormal
    
    '---------------------------------------------------
    'TODO everywhere when choosen show print in a window
    Xlog.WindowState = vbNormal
    Xlog.statusForm = ""
    Xlog.Width = 11595
    '---------------------------------------------------
        
    Xlog.Afsluiten.Enabled = False
    Xlog.WijzigenLijn.Visible = False
    Xlog.Afsluiten.TabStop = False
    Xlog.cbAfbeelding.Visible = False
    XLogKey = ""
    Xlog.SSTab1.TabVisible(1) = False
    Xlog.Show 1
    Unload Xlog
End If

End Sub



Private Sub chkAfdrukLiggend_Click()

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If chkAfdrukLiggend.Value = 1 Then
    Printer.Orientation = vbPRORLandscape
Else
    Printer.Orientation = vbPRORPortrait
End If

End Sub

Private Sub Drukken_Click()
Dim BeginSleutel    As String
Dim EindSleutel     As String
Dim bModDummy       As Integer
Dim Lijn            As Integer
Dim DCBedrag        As Currency
Dim XX As Integer

aa = ""
TLijnen = 0
Lijn = 0
BeginSleutel = "D0" + PlGrensVan
EindSleutel = "D0" + PlGrensTot

    psTekst(2) = "Diverse Postenboek " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
    psTekst(0) = TekstLijn(1).text
    psTekst(3) = TekstLijn(0).text

InitVelden
bGetOrGreater TABLE_JOURNAL, 1, Trim$(BeginSleutel)
If Ktrl Or KEY_BUF(TABLE_JOURNAL) > EindSleutel Then
    Beep
    Exit Sub
Else
    Screen.MousePointer = vbHourglass
    bClose TABLE_DUMMY
    ClearFlDummy
    Ktrl = bOpen(TABLE_DUMMY)
    
    PAGE_COUNTER = 0
    If chkAfdrukInVenster = 0 Then
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
            Printer.FontBold = True
        End If
    End If
    PrintTitel
    GoSub PrintInfo
End If

Do
    bNext TABLE_JOURNAL
    If Ktrl Or KEY_BUF(TABLE_JOURNAL) > EindSleutel Then
        Exit Do
    Else
        GoSub PrintInfo
    End If
Loop
PrintTotaal
CumulPrint
Printer.EndDoc
Screen.MousePointer = vbNormal
Annuleren_Click
Exit Sub

PrintInfo:
RecordToVeld TABLE_JOURNAL
Lijn = Lijn + 1
VeldTXT(0) = Format(Lijn, "0000")
VeldTXT(1) = DATE_TEXT(objectValue(rsMAR(TABLE_JOURNAL)("v066").Value))
VeldTXT(2) = vBibTekst(TABLE_JOURNAL, "#v019 #")
bGet TABLE_LEDGERACCOUNTS, 0, vSet(VeldTXT(2), 7)
If Ktrl Then
    VeldTXT(3) = "Reeds vernietigd..."
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    VeldTXT(3) = Left(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), 36)
End If
VeldTXT(4) = objectValue(rsMAR(TABLE_JOURNAL)("v067").Value)
DCBedrag = objectValue(rsMAR(TABLE_JOURNAL)("dece068").Value)
SnelHelpPrint VeldTXT(1) + " " + VeldTXT(4), BL_LOGGING
Select Case DCBedrag
    Case Is < 0
        TotaalC = TotaalC + DCBedrag
        VeldTXT(5) = ""
        VeldTXT(6) = Dec$(Abs(DCBedrag), MASK_EURBH)
    Case Else
        TotaalD = TotaalD + DCBedrag
        VeldTXT(5) = Dec$((DCBedrag), MASK_EURBH)
        VeldTXT(6) = ""
End Select
VeldTXT(7) = vBibTekst(TABLE_JOURNAL, "#v069 #")
GoSub DetailCumul
PrintVelden
Return

PrtHandler9:
MsgBox "Kontroleer de printer."
Resume

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

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Dim T As Integer
Dim TempB_MODUS As Integer

TotaalD = 0
TotaalC = 0

TekstLijn(0).text = BYPERDAT.PeriodeBoekjaar.text
TekstLijn(1).text = MIM_GLOBAL_DATE

PlGrensVan = Left(PERIOD_FROMTO, 8)
PlGrensTot = Right(PERIOD_FROMTO, 8)

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If Printer.Orientation = vbPRORLandscape Then
    chkAfdrukLiggend.Value = 1
Else
    chkAfdrukLiggend.Value = 0
End If
chkAfdrukLiggend_Click

End Sub

Private Sub InitVelden()
Dim T As Integer

REPORT_FIELD(0) = "Lijn"
    REPORT_TAB(0) = 2

REPORT_FIELD(1) = "Datum"
    REPORT_TAB(1) = 7

REPORT_FIELD(2) = "Nummer"
    REPORT_TAB(2) = 18

REPORT_FIELD(3) = "Naam Rekening"
    REPORT_TAB(3) = 26

REPORT_FIELD(4) = "Betreft"
    REPORT_TAB(4) = 63

REPORT_FIELD(5) = "       Debet"
    REPORT_TAB(5) = 94
    
REPORT_FIELD(6) = "      Credit"
    REPORT_TAB(6) = 106

REPORT_FIELD(7) = "T.Rekening"
    REPORT_TAB(7) = 119

REPORT_TAB(8) = 0

If chkAfdrukInVenster Then
    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = "Diverse postenboek"
    Xlog.X.Cols = 8
    Xlog.X.Row = 0
    For T = 0 To 7
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

Printer.Print vbCrLf$; FULL_LINE

Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print REPORT_FIELD(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop

Printer.Print FULL_LINE; vbCrLf$; vbCrLf$;

Exit Sub

PrtHandler3:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub PrintTotaal()
Dim T As Integer

For T = 0 To 7
    VeldTXT(T) = ""
Next
VeldTXT(4) = "Periodiek totaal :"
VeldTXT(5) = Dec$(Abs(TotaalD), MASK_EURBH)
VeldTXT(6) = Dec$(Abs(TotaalC), MASK_EURBH)

If chkAfdrukInVenster Then
Else
    Printer.Print vbCrLf; FULL_LINE
End If

T = 0
aa = ""
Do While T < 8
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print ; Tab(REPORT_TAB(T));
        Printer.Print ; VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
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
Dim aa As String

T = 0
aa = ""
Do While REPORT_TAB(T) <> 0
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print Tab(REPORT_TAB(T));
        Printer.Print VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If Printer.CurrentY >= Printer.Height - 1200 Then
    On Local Error Resume Next
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If

If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1

End Sub


Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelLength = Len(TekstLijn(Index).text)

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 0
        If DATE_INVALID(Right(TekstLijn(0).text, 10)) Then
            MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
            TekstLijn(0).text = BYPERDAT.PeriodeBoekjaar.text
            Exit Sub
        ElseIf Len(TekstLijn(0).text) <> 23 Then
            MsgBox "Respecteer : " + vbCrLf + vbCrLf + "DD/MM/EEJJ - DD/MM/EEJJ a.u.b. !"
            TekstLijn(0).text = BYPERDAT.PeriodeBoekjaar.text
            Exit Sub
        Else
            PlGrensVan = Mid(TekstLijn(0).text, 7, 4) + Mid(TekstLijn(0).text, 4, 2) + Mid(TekstLijn(0).text, 1, 2)
            PlGrensTot = Mid(TekstLijn(0).text, 20, 4) + Mid(TekstLijn(0).text, 17, 2) + Mid(TekstLijn(0).text, 14, 2)
        End If
    Case 1
        If DATE_INVALID((TekstLijn(1).text)) Then
            Beep
            TekstLijn(1).text = MIM_GLOBAL_DATE
        End If
    Case 3
        TekstLijn(3).text = Format(Val(TekstLijn(3).text), "00000")
End Select

End Sub

