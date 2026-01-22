VERSION 5.00
Begin VB.Form Rapportage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rapportage boekhouding : "
   ClientHeight    =   3195
   ClientLeft      =   1650
   ClientTop       =   1155
   ClientWidth     =   6180
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
   ScaleHeight     =   3195
   ScaleWidth      =   6180
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
      Left            =   4680
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   1400
   End
   Begin VB.ComboBox Keuzelijst 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   5955
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&drukken"
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
      Left            =   3240
      TabIndex        =   2
      Top             =   60
      Width           =   1400
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1020
      TabIndex        =   0
      Top             =   60
      Width           =   1800
   End
   Begin VB.Label LabelMemo 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   60
      TabIndex        =   5
      Top             =   960
      Width           =   6015
   End
   Begin VB.Label Label1 
      Caption         =   "Datum Drukken"
      Height          =   375
      Index           =   1
      Left            =   180
      TabIndex        =   4
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "Rapportage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim psTekst(5) As String        'Koptekstinfo

Dim sFile(20) As String
Dim smin(10) As String
Dim smax(10) As String
Dim smg(10) As String
Dim stekst(10) As String

Dim ichr(10)                As Integer
Dim LFontSize(20)           As Single
Dim LAantalL(20)            As Integer
Dim FontDefChanged          As Integer

Private Sub Annuleren_Click()

Unload Me

End Sub

Private Sub Drukken_Click()
Dim FlTemp As Integer
Dim ideler As Integer
Dim lijntel As Integer

Dim dbdhbj As Currency
Dim dbdvbj As Currency
Dim dTOT As Currency
Dim dvtot As Currency

'nieuw rekenkundige controle
Dim rTotaal1 As Currency
Dim rTotaal2 As Currency
'einde

Dim sminsl As String * 7
Dim smaxsl As String * 7

Dim sbd As String
Dim alijn As Integer
Dim Scode As String
Dim apos As Integer

LabelMemo.Caption = ""
Screen.MousePointer = vbHourglass

On Local Error GoTo PrtHandler2

    PAGE_COUNTER = 0
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
            Printer.FontBold = True
        End If
    
FlTemp = FreeFile
Open PROGRAM_LOCATION + "Def\" + sFile(Keuzelijst.ListIndex) For Input As FlTemp
    Input #FlTemp, smg(9)
    Input #FlTemp, smg(2)
    Input #FlTemp, smg(3)
    Input #FlTemp, smg(4)
    Input #FlTemp, smg(5)
    Input #FlTemp, ideler

lijntel = 6
If ideler = 1 Then
    ideler = 1000
Else
    ideler = 1
End If

psTekst(2) = "Rapportage " + UCase(Mid(Mim.Caption, InStr(Mim.Caption, "[")))
psTekst(0) = TekstLijn(1).Text
psTekst(3) = smg(9) + " Boekjaar " + BJPERDAT.Boekjaar.Text

While Not EOF(FlTemp)
    lijntel = lijntel + 1
    Input #FlTemp, sbd, alijn, Scode, apos
    If InStr("DCXYmheltr", sbd) = 0 Then
        'nieuw rekenkundige controle
        MsgBox "Rapportagedefinitiebestand defekt of beschadigd."
        Close FlTemp
        Printer.NewPage
        Printer.EndDoc
        Exit Sub
    End If
        
    For T = 0 To alijn - 1
        lijntel = lijntel + 1
        Input #FlTemp, stekst(T)
        SnelHelpPrint "LN:" + Dec$((lijntel), "###") + "  " + stekst(T), BL_LOGGING
    Next T

    lijntel = lijntel + 1
    Input #FlTemp, asort
    For T = 1 To asort
        lijntel = lijntel + 1
        Input #FlTemp, smin(T), smax(T)
    Next T

    If InStr("mhelt", sbd) Then
        GoSub Label1000
    ElseIf InStr("r", sbd) Then
        'nieuw rekenkundige controle
        GoSub LabelRekening
    Else
        For klt = 1 To asort - 1
            cflag = 1
            sminsl = smin(klt)
            smaxsl = smax(klt)
            GoSub Label1000
        Next klt
        sminsl = smin(asort)
        smaxsl = smax(asort)
        cflag = 0
        GoSub Label1000
    End If
Wend
Close FlTemp
Printer.EndDoc
Screen.MousePointer = vbNormal
Exit Sub

'nieuw rekenkundige controle
LabelRekening:
rTotaal1 = 0
rTotaal2 = 0
For klt = 1 To asort
    cflag = 1
    dTOT = 0: dvtot = 0
    sminsl = Mid(smin(klt), 2, 7)
    smaxsl = Mid(smin(klt), 10, 7)
    GoSub Label1000
    If Left(smin(klt), 1) = "+" Then
        rTotaal1 = rTotaal1 + dTOT
    Else
        'Stop
    End If
    dTOT = 0: dvtot = 0
    sminsl = Mid(smax(klt), 2, 7)
    smaxsl = Mid(smax(klt), 10, 7)
    GoSub Label1000
    If Left(smax(klt), 1) = "+" Then
        rTotaal2 = rTotaal2 + dTOT
    Else
        'Stop
    End If
Next klt
Select Case Left(Scode, 2)
    Case "? "
        If rTotaal1 = 0 Or rTotaal2 = 0 Then
        Else
            Msg = Mid(Mim.SnelHelp.SimpleText, 9) + vbCrLf + vbCrLf
            Msg = Msg + Format(rTotaal1, MASK_EURBH) + " ? " + Format(rTotaal2, MASK_EURBH)
            MsgBox Msg, 0, "Kontroleer a.u.b. !"
        End If
    Case Else
        'Stop
End Select
Return

Label1000:
Select Case sbd
    Case "m"
        Tekst$ = ""
        For T = 0 To alijn
            Tekst$ = Tekst$ + stekst(T) + vbCrLf
        Next
        LabelMemo.Caption = Tekst$
        LabelMemo.Refresh
        GoTo Label1060

    Case "h"
        GoSub Label2000
        GoTo Label1060

    Case "e"
        GoTo Label2200

    Case "t"
        GoTo Label2300

    Case "l"
        GoTo Label2400
    Case Else
End Select

Vooraan:
    bFirst TABLE_LEDGERACCOUNTS, 0
    bGetOrGreater TABLE_LEDGERACCOUNTS, 0, sminsl
    If Ktrl Then
        GoTo PrintAf
    Else
        If vSet(KEY_BUF(TABLE_LEDGERACCOUNTS), 7) > smaxsl Then
            GoTo PrintAf
        Else
            RecordToVeld TABLE_LEDGERACCOUNTS
            GoSub TelOp
        End If
    End If

Do
    bNext TABLE_LEDGERACCOUNTS
    If Ktrl Then
        GoTo PrintAf
    Else
        If vSet(KEY_BUF(TABLE_LEDGERACCOUNTS), 7) > smaxsl Then
            GoTo PrintAf
        Else
            RecordToVeld TABLE_LEDGERACCOUNTS
            GoSub TelOp
        End If
    End If
Loop
MsgBox "stop na DO-LOOP"

TelOp:
If bhEuro Then
    dbdhbj = Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(22 + BJPERDAT.Boekjaar.ListIndex, "000") + " #"))
    dbdvbj = Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(22 + BJPERDAT.Boekjaar.ListIndex + 1, "000") + " #"))
Else
    dbdhbj = Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(22 + BJPERDAT.Boekjaar.ListIndex, "000") + " #"))
    dbdvbj = Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(22 + BJPERDAT.Boekjaar.ListIndex + 1, "000") + " #"))
End If
dTOT = dTOT + dbdhbj
dvtot = dvtot + dbdvbj
Return

PrintAf:
alt = 0
While stekst(alt) <> ""
    Printer.Print Tab(1); stekst(alt); " "; Chr$(ichr(5));
    If stekst(alt + 1) <> "" Then
        Printer.Print Space$(7); Chr$(ichr(5)); Space$(23); Chr$(ichr(5)); Space$(24); Chr$(ichr(5))
    End If
    alt = alt + 1
Wend

If cflag = 1 Then
    GoTo Label1060
ElseIf sbd = "C" Then
    dTOT = -dTOT
    dvtot = -dvtot
End If

If sbd = "X" Then
    If dvtot > 0 Then
        dvtot = 0
    Else
        dvtot = Abs(dvtot)
    End If
End If

If sbd = "Y" Then
    If dvtot < 0 Then
        dvtot = 0
    End If
End If

If sbd = "X" Then
    If dTOT > 0 Then
        dTOT = 0
    Else
        dTOT = Abs(dTOT)
    End If
End If

If sbd = "Y" Then
    If dTOT < 0 Then
        dTOT = 0
    End If
End If

Printer.Print Scode; Chr$(ichr(5)); Tab(72 + apos - 2); Dec$(dTOT / ideler, MASK_EURBH);
Printer.Print Tab(93); Chr$(ichr(5)); Tab(95 + apos); Dec$(dvtot / ideler, MASK_EURBH);
Printer.Print Tab(117); " "; Chr$(ichr(5))
dTOT = 0
dvtot = 0

Label1060:
For tlp = 0 To 5
    stekst(tlp) = ""
Next tlp
Return

Label2000:
PrintTitel
Printer.Print String$(Len(stekst(0)) + 1, Chr$(ichr(4))); Chr$(ivbtab); String$(7, Chr$(ichr(4))); Chr$(ivbtab); String$(23, Chr$(ichr(4))); Chr$(ivbtab); String$(24, Chr$(ichr(4))); Chr$(ichr(1))
Printer.Print Space$(Len(stekst(0)) + 1); Chr$(ichr(5)); Space$(7); Chr$(ichr(5)); smg(2); Chr$(ichr(5)); smg(3); Chr$(ichr(5))
Printer.Print Space$(Len(stekst(0)) + 1); Chr$(ichr(5)); smg(4); Chr$(ichr(6)); String$(23, Chr$(ichr(4))); Chr$(ichr(8)); String$(24, Chr$(ichr(4))); Chr$(ichr(7))
Printer.Print Space$(Len(stekst(0)) + 1); Chr$(ichr(5)); Space$(7); Chr$(ichr(5)); smg(5); Chr$(ichr(5))
Printer.Print Space$(Len(stekst(0)) + 1); Chr$(ichr(6)); String$(7, Chr$(ichr(4))); Chr$(ichr(10)); String$(23, Chr$(ichr(4))); Chr$(ivbtab); String$(24, Chr$(ichr(4))); Chr$(ichr(7))
alt = 0
While stekst(alt) <> ""
    Printer.Print Tab(1); stekst(alt); " "; Chr$(ichr(5)); Space$(7); Chr$(ichr(5)); Space$(23); Chr$(ichr(5)); Space$(24); Chr$(ichr(5))
    alt = alt + 1
Wend
Return

Label2200:
Printer.Print Space$(60); Chr$(ichr(2)); String$(7, Chr$(ichr(4))); Chr$(ichr(8)); String$(23, Chr$(ichr(4))); Chr$(ichr(8)); String$(24, Chr$(ichr(4))); Chr$(ichr(3));
    Printer.NewPage
Return

Label2300:
Printer.Print Space$(60); Chr$(ichr(6)); String$(7, Chr$(ichr(4))); Chr$(ichr(10)); String$(23, Chr$(ichr(4))); Chr$(ichr(10)); String$(24, Chr$(ichr(4))); Chr$(ichr(7))
Return

Label2400:
Printer.Print Space$(60); Chr$(ichr(5)); Space$(7); Chr$(ichr(5)); Space$(23); Chr$(ichr(5)); Space$(24); Chr$(ichr(5))
Return

PrtHandler2:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Dim Bedrijf As String
Dim sInfo As String
Dim FlTemp As Integer
Dim i As Integer
Dim T As Integer
Dim TempB_MODUS As Integer

Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2

TekstLijn(1).Text = MIM_GLOBAL_DATE
Caption = Caption + " " + BJPERDAT.Boekjaar.Text

i = 0
Bedrijf = Dir$(PROGRAM_LOCATION + "Def\" + "99MAR*.SRN")
If Bedrijf = Space$(Len(Bedrijf)) Then
    Drukken.Enabled = False
Else
    Do While Bedrijf <> ""
        FlTemp = FreeFile
        Open PROGRAM_LOCATION + "Def\" + Bedrijf For Input As FlTemp
            Input #FlTemp, sInfo
        Close FlTemp
        sInfo = UCase$(Left(sInfo, 1)) + LCase$(Right(sInfo, Len(sInfo) - 1))
        sFile(i) = Bedrijf
        Keuzelijst.AddItem sInfo
        i = i + 1
        Bedrijf = Dir$
    Loop
    If Keuzelijst.ListCount Then
        Keuzelijst.ListIndex = i - 1
    End If
End If

For i = 0 To 10
    ichr(i) = 124
Next

End Sub

Private Sub Keuzelijst_Click()

LabelMemo.Caption = ""

End Sub

Private Sub PrintTitel()
Dim T As Integer

On Local Error GoTo PrtHandler1

If usrLicentieInfo <> "" Then
    Printer.CurrentX = 50
    Printer.CurrentY = 50
    Printer.Print usrLicentieInfo;
End If
PAGE_COUNTER = PAGE_COUNTER + 1
Printer.CurrentY = 400
Printer.Print Tab(1); psTekst(2); Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########");
Printer.Print Tab(108); "Datum  : "; psTekst(0); vbCrLf$; vbCrLf$;
Printer.Print Tab(1); UCase$(psTekst(3));

Printer.Print vbCrLf; FULL_LINE; vbCrLf; vbCrLf;
Exit Sub

PrtHandler1:
MsgBox "Kontroleer de printer."
Resume

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 1
        If DATE_INVALID((TekstLijn(1).Text)) Then
            Beep
            TekstLijn(1).Text = MIM_GLOBAL_DATE
            TekstLijn(1).SetFocus
        End If
End Select

End Sub

