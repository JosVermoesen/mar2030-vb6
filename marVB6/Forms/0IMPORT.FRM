VERSION 5.00
Begin VB.Form Importeren 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bestanden Importeren"
   ClientHeight    =   1545
   ClientLeft      =   1515
   ClientTop       =   1845
   ClientWidth     =   5430
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1545
   ScaleWidth      =   5430
   Begin VB.CommandButton DefEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TabStart"
      Enabled         =   0   'False
      Height          =   405
      Index           =   3
      Left            =   4110
      TabIndex        =   6
      Top             =   3750
      Width           =   1245
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   5400
      TabIndex        =   7
      Top             =   4560
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   5400
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.ComboBox Formattering 
      Appearance      =   0  'Flat
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
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4560
      Visible         =   0   'False
      Width           =   4035
   End
   Begin VB.CommandButton DefEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Formattering"
      Enabled         =   0   'False
      Height          =   405
      Index           =   2
      Left            =   4110
      TabIndex        =   17
      Top             =   3300
      Width           =   1245
   End
   Begin VB.CommandButton DefEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Interpreteer"
      Enabled         =   0   'False
      Height          =   405
      Index           =   1
      Left            =   4110
      TabIndex        =   16
      Top             =   2850
      Width           =   1245
   End
   Begin VB.ListBox RapportVelden 
      Appearance      =   0  'Flat
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
      Height          =   2130
      Left            =   5400
      TabIndex        =   15
      Top             =   2400
      Width           =   4035
   End
   Begin VB.CommandButton DefEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Toevoegen"
      Enabled         =   0   'False
      Height          =   405
      Index           =   0
      Left            =   4110
      TabIndex        =   14
      Top             =   2400
      Width           =   1245
   End
   Begin VB.ListBox TabelVelden 
      Appearance      =   0  'Flat
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
      Height          =   2130
      Left            =   90
      MultiSelect     =   2  'Extended
      TabIndex        =   12
      Top             =   2400
      Width           =   3975
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   1530
      TabIndex        =   10
      Top             =   1740
      Width           =   5925
   End
   Begin VB.CommandButton Sluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   405
      Left            =   3720
      TabIndex        =   9
      Top             =   990
      Width           =   1545
   End
   Begin VB.CommandButton Start 
      Caption         =   "1, 2, 3,... Start!"
      Height          =   405
      Left            =   3720
      TabIndex        =   8
      Top             =   540
      Width           =   1545
   End
   Begin VB.CommandButton Definitie 
      Caption         =   "Andere definitie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3720
      TabIndex        =   4
      Top             =   90
      Width           =   1545
   End
   Begin VB.ComboBox BestandsDefinitie 
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
      Left            =   90
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   3285
   End
   Begin VB.ComboBox Tabel 
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
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   2415
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Rapportvelden"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   5430
      TabIndex        =   18
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tabelvelden"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   13
      Top             =   2190
      Width           =   1185
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Rapportnaam"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   5
      Left            =   150
      TabIndex        =   11
      Top             =   1830
      Width           =   1275
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   0
      X2              =   9570
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      Caption         =   "Aktieve bestandsdefinitie"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Tabel"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   645
   End
End
Attribute VB_Name = "Importeren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim vBibDef As String * 5
Dim TabDef As Integer
Dim PosLen As Integer
Dim VeldInfo As String
Dim RecordIMP As String
Dim FlHier As Integer

Private Sub BestandsDefinitie_Click()
    
    Dim FlTemp As Integer
    FlTemp = FreeFile
    Open PROGRAM_LOCATION + "Def\" + Format(Val(Left(Tabel.Text, 1)), "000") + Left(BestandsDefinitie.Text, 2) + ".IOD" For Input As FlTemp
    Line Input #FlTemp, RapportTekst$
    TekstInfo(2).Text = RapportTekst$
    RapportVelden.Clear
    Do While Not EOF(FlTemp)
        Line Input #FlTemp, RapportTekst$
        RapportVelden.AddItem RapportTekst$
    Loop
    Close FlTemp

End Sub

Private Sub BestandsDefinitie_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 46
        Msg = "BestandsDefinitie " + vbCrLf + vbCrLf + BestandsDefinitie.Text + vbCrLf + vbCrLf + "verwijderen.  Bent U zeker ?"
        KtrlBox = MsgBox(Msg, 292)
        If KtrlBox = 6 Then
            Kill PROGRAM_LOCATION + "Def\" + Format(Val(Left(Tabel.Text, 1)), "000") + Left(BestandsDefinitie.Text, 2) + ".IOD"
            BestandsDefinitie.RemoveItem BestandsDefinitie.ListIndex
        End If
End Select

End Sub

Private Sub DefEdit_Click(Index As Integer)

Select Case Index
    Case 0
        For i = 0 To TabelVelden.ListCount - 1
            If TabelVelden.Selected(i) Then
                TempoString$ = TabelVelden.List(i)
                If TabLijn + Val(Mid(TempoString$, 16, 3)) > 128 Then
                    TabLijn = 1
                Else
                    TabLijn = TabLijn + 1
                End If
                Mid(TempoString$, 12, 3) = Format(TabLijn, "000")
                RapportVelden.AddItem TempoString$
                TabLijn = TabLijn + Val(Mid(TempoString$, 16, 3))
            End If
        Next i
    Case 1
        If RapportVelden.ListIndex < 0 Then Beep: Exit Sub
        TekstInfo(3).Text = Mid(RapportVelden.Text, 22)
        TekstInfo(3).Visible = True
        TekstInfo(3).SetFocus

    Case 2
        If RapportVelden.ListIndex < 0 Then Beep: Exit Sub
        For T = 0 To Formattering.ListCount - 1
            Formattering.ListIndex = T
            If Mid(RapportVelden.Text, 20, 1) = Left(Formattering.Text, 1) Then
                Exit For
            End If
        Next
        Formattering.Visible = True
        Formattering.SetFocus
    
    Case 3
        If RapportVelden.ListIndex < 0 Then Beep: Exit Sub
        TekstInfo(4).Text = Mid(RapportVelden.Text, 12, 3)
        TekstInfo(4).Visible = True
        TekstInfo(4).SetFocus

End Select

End Sub

Private Sub Definitie_Click()

If Start.Enabled = True Then
    If TLBPag2(Format(Val(Left(Tabel.Text, 1)), "000")) Then
        T = 0
        TabelVelden.Clear
        Do While TELEBIB_CODE(T) <> Space$(10)
            TabelVelden.AddItem TELEBIB_CODE(T) + " " + "000 " + Format(TELEBIB_LENGTH(T), "000") + " T " + RTrim$(TELEBIB_TEXT(T))
            T = T + 1
        Loop
    Else
        Beep
        Exit Sub
    End If
    Start.Enabled = False
    TabelVelden.Enabled = True
    RapportVelden.Enabled = True
    TekstInfo(2).Enabled = True
    Tabel.Enabled = False
    BestandsDefinitie.Enabled = False
    DefEdit(0).Enabled = True
    WindowState = 2
    TekstInfo(2).SetFocus
Else
    Start.Enabled = True
    TabelVelden.Enabled = False
    RapportVelden.Enabled = False
    TekstInfo(2).Enabled = False
    Tabel.Enabled = True
    BestandsDefinitie.Enabled = True
    For T = 0 To 3
        DefEdit(T).Enabled = False
    Next
    WindowState = 0
    Dim FlTemp As Integer
    FlTemp = FreeFile
    VolgNummer% = BestandsDefinitie.ListCount
    Open PROGRAM_LOCATION + "Def\" + Format(Val(Left(Tabel.Text, 1)), "000") + Format(VolgNummer%, "00") + ".IOD" For Output As FlTemp
    Print #FlTemp, (TekstInfo(2).Text)
    For T = 0 To RapportVelden.ListCount - 1
        Print #FlTemp, (RapportVelden.List(T))
    Next
    Close FlTemp
    BestandsDefinitie.AddItem Format(VolgNummer%, "00") + ": " + TekstInfo(2).Text
    BestandsDefinitie.ListIndex = BestandsDefinitie.ListCount - 1
End If

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Left = 0
Top = 0

Tabel.AddItem "1: Klanten"
Tabel.AddItem "2: Leveranciers"
Tabel.AddItem "3: Algemene rekeningen"
Tabel.AddItem "4: Produkten"

Formattering.AddItem "T: Tekst zonder enige bewerking"
Formattering.AddItem "D: Van SorteerDatum naar DD/MM/EEJJ"
Formattering.AddItem "0: Bedrag met masker #########"
Formattering.AddItem "1: Bedrag met masker ####"
Formattering.AddItem "2: Bedrag met masker #######.##"
Formattering.AddItem "3: Bedrag met masker ###.########"
Formattering.AddItem "4: Bedrag met masker ########.##"
Formattering.AddItem "5: Bedrag met masker ###"
Formattering.AddItem "6: Bedrag met masker ##"
Tabel.ListIndex = 0

End Sub

Private Sub Formattering_LostFocus()

Positie = RapportVelden.ListIndex
TempoString$ = RapportVelden.Text
Mid(TempoString$, 20, 1) = Left(Formattering.Text, 1)
RapportVelden.RemoveItem Positie
RapportVelden.AddItem TempoString$, Positie
RapportVelden.SetFocus
RapportVelden.ListIndex = Positie
Formattering.Visible = False

End Sub


Private Sub ImportVelden()
Dim T As Integer
Static OptelLijn As Integer

On Local Error GoTo ErrorImport
For T = 0 To RapportVelden.ListCount - 1
    vBibDef = Mid(RapportVelden.List(T), 5, 5)
    If Mid(RapportVelden.List(T), 22, 1) = "=" Then
        VeldInfo = Mid(RapportVelden.List(T), 23, InStr(23, RapportVelden.List(T), ";") - 23)
    Else
        TabDef = Val(Mid(RapportVelden.List(T), 12, 3))
        PosLen = Val(Mid(RapportVelden.List(T), 22, InStr(22, RapportVelden.List(T), ";")))
        VeldInfo = SuperCase(Mid(RecordIMP, TabDef, PosLen))
    End If
    If T = 1 Then
        OptelLijn = OptelLijn + 1
        SnelHelpPrint Format(OptelLijn) + " " + VeldInfo, BL_LOGGING
    End If
    If Trim$(vBibDef) = "v112" Then
        If InStr(VeldInfo, ".") Then
            VeldInfo = Left(VeldInfo, InStr(VeldInfo, ".") - 1) + Mid(VeldInfo, InStr(VeldInfo, ".") + 1)
        End If
    End If
    vBib FlHier, VeldInfo, vBibDef    'Taalkode
Next

ErrorImport:
Exit Sub

End Sub

Private Sub NieuweFiche(Fl As Integer)

TLB_RECORD(Fl) = ""

Select Case Fl
    Case TABLE_CUSTOMERS, TABLE_SUPPLIERS
        vBib Fl, "2", "A10C"    'Taalkode
        vBib Fl, "002", "v149"  'Landnummer  ISO kode
        vBib Fl, "B  ", "A109"  'Landkode Postkantoor
        vBib Fl, "BE", "v150"   'Landkode    ISO kode
        vBib Fl, "BEF", "vs03"  'Munteenheid ISO kode
        vBib Fl, "1", "vs07"    'exemplaren dokumenten
    Case TABLE_LEDGERACCOUNTS
        vBib Fl, "O", "v032"    'Budgetcode
    Case TABLE_PRODUCTS
        vBib Fl, "0", "v106"    'ProduktMaatstaf
        vBib Fl, "1", "v107"    'AantalMaatstaven in verpakking
        vBib Fl, "N", "v108"    'Soort produkt
        vBib Fl, "3", "v111"    'Btw Kode
        vBib Fl, String99(READING, 77), "v116"  'aankooprekening
        vBib Fl, String99(READING, 78), "v117"  'verkooprekening
        vBib Fl, String99(READING, 79), "v118"  'voorraadrekening
    Case Else
        MsgBox "stop"
End Select

End Sub

Private Sub RapportVelden_GotFocus()

DefEdit(0).Enabled = False
If RapportVelden.ListCount = 0 Then
    TabLijn = 0
Else
    DefEdit(1).Enabled = True
    DefEdit(2).Enabled = True
    DefEdit(3).Enabled = True
End If

End Sub

Private Sub RapportVelden_KeyDown(KeyCode As Integer, Shift As Integer)

If RapportVelden.ListCount = 0 Then TabLijn = 0: Exit Sub

Select Case KeyCode
    Case 46
        Positie = RapportVelden.ListIndex
        If Positie >= 0 Then
            RapportVelden.RemoveItem Positie
        End If
End Select

End Sub

Private Sub sluiten_Click()

WindowState = 0
Unload Me

End Sub

Private Sub Start_Click()
Dim FlTemp As Integer
Dim ImportBestand As String

On Local Error GoTo ErrorJump
    
FlHier = Val(Left(Tabel.Text, 1))
bClose FlHier

    Mim.Teken.filename = "*.*"
    Mim.Teken.Filter = "Enkel Tekstbestanden a.u.b. !!! (*.*)|*.*"
    Mim.Teken.CancelError = True
    Mim.Teken.Action = 1
    ImportBestand = Mim.Teken.filename + " "
    
On Local Error GoTo 0
Msg = "Gegevens uit " + ImportBestand
Msg = Msg + " invoegen bij " + Mid(Tabel.Text, 3)
Msg = Msg + vbCrLf + vbCrLf + "Bent U zeker ?"
KtrlBox = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2)
If KtrlBox = vbYes Then
Else
    Exit Sub
End If

FlTemp = FreeFile
Open ImportBestand For Input As FlTemp
Do While Not EOF(FlTemp)
    Line Input #FlTemp, RecordIMP
    vBibDef = Mid(RapportVelden.List(0), 5, 5)
    If Mid(RapportVelden.List(0), 22, 1) = "=" Then
        GoTo ErrorJump
    Else
        TabDef = Val(Mid(RapportVelden.List(0), 12, 3))
        PosLen = Val(Mid(RapportVelden.List(0), 22, InStr(22, RapportVelden.List(0), ";")))
        VeldInfo = Mid(RecordIMP, TabDef, PosLen)
        bGet FlHier, 0, vSet(VeldInfo, FLINDEX_LEN(FlHier, 0))
        If Ktrl Then
            NieuweFiche FlHier
            ImportVelden
            bInsert FlHier, 0
        Else
            RecordToVeld FlHier
            ImportVelden
            bUpdate FlHier, 0
        End If
    End If
Loop
Close FlTemp
bClose FlHier
MsgBox "Import succesvol beëindigd"
Exit Sub

ErrorJump:
bClose FlHier
MsgBox "Import NIET succesvol beëindigd !"
Exit Sub
Resume

End Sub

Private Function SuperCase(Wijzigen As String) As String

On Local Error GoTo SuperError

SuperCase = UCase$(Left(Wijzigen, 1)) + LCase$(Right(Wijzigen, Len(Wijzigen) - 1))
Exit Function

SuperError:
Exit Function

End Function

Private Sub Tabel_Click()
Dim T As Integer
Dim FlL As Integer

BestandsDefinitie.Clear
X$ = Dir$(PROGRAM_LOCATION + "Def\" + Format(Val(Left(Tabel.Text, 1)), "000") + "??.IOD")
If X$ = "" Then
    Start.Enabled = False
    RapportVelden.Clear
    TekstInfo(2).Text = ""
    Exit Sub
Else
    Start.Enabled = True
    FlL = FreeFile
    Do While X$ <> ""
        Open PROGRAM_LOCATION + "Def\" + X$ For Input As FlL

            Line Input #FlL, TekstLijn$
        Close FlL
        BestandsDefinitie.AddItem Mid(X$, 4, 2) + ": " + TekstLijn$
        X$ = Dir$
    Loop
End If

If BestandsDefinitie.ListCount <> 0 Then
    BestandsDefinitie.ListIndex = 0
End If

End Sub

Private Sub TabelVelden_Click()

SnelHelpPrint (TabelVelden.Text), BL_LOGGING

End Sub

Private Sub TabelVelden_GotFocus()

DefEdit(0).Enabled = True
DefEdit(1).Enabled = False
DefEdit(2).Enabled = False
DefEdit(3).Enabled = False

End Sub

Private Sub TekstInfo_GotFocus(Index As Integer)

TekstInfo(Index).SelLength = Len(TekstInfo(Index).Text)

End Sub

Private Sub TekstInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case Index
    Case 3
        Select Case KeyCode
            Case 13
                Positie = RapportVelden.ListIndex
                TempoString$ = Left(RapportVelden.Text, 21) + TekstInfo(3).Text
                RapportVelden.RemoveItem Positie
                RapportVelden.AddItem TempoString$, Positie
                RapportVelden.SetFocus
                RapportVelden.ListIndex = Positie
            End Select
    Case 4
        Select Case KeyCode
            Case 13
                Positie = RapportVelden.ListIndex
                TempoString$ = RapportVelden.Text
                Mid(TempoString$, 12, 3) = Format(Val(TekstInfo(4).Text), "000")
                RapportVelden.RemoveItem Positie
                RapportVelden.AddItem TempoString$, Positie
                RapportVelden.SetFocus
                RapportVelden.ListIndex = Positie
            End Select
End Select

End Sub

Private Sub TekstInfo_LostFocus(Index As Integer)

Select Case Index
    Case 3, 4
        TekstInfo(Index).Visible = False
    Case 5
        If Val(TekstInfo(5).Text) < 1 Or Val(TekstInfo(5).Text) > FLINDEX_LEN(FlKeuze, IndexKeuze) Then
            Beep
            TekstInfo(Index).Text = Str$(FLINDEX_LEN(FlKeuze, IndexKeuze))
            TekstInfo(Index).SetFocus
        End If
End Select

End Sub

