VERSION 5.00
Begin VB.Form SQLLijsten 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ANSI-92 SQL Lijst rapportage"
   ClientHeight    =   2025
   ClientLeft      =   1845
   ClientTop       =   2475
   ClientWidth     =   7410
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2025
   ScaleWidth      =   7410
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSQLOvername 
      Caption         =   "SQL Resultaat via Generator"
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
      Left            =   4890
      TabIndex        =   32
      Top             =   930
      Width           =   2475
   End
   Begin VB.CommandButton cmdToonSQL 
      Caption         =   "&Toon SQL SELECT Definitie"
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
      Left            =   4890
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   600
      Width           =   2475
   End
   Begin VB.CheckBox chkAfdrukLiggend 
      Caption         =   "Liggende Printerafdruk"
      Height          =   255
      Left            =   60
      TabIndex        =   29
      Top             =   1620
      Width           =   2310
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Alignment       =   1  'Right Justify
      Caption         =   "Afdruk in venster"
      Height          =   255
      Left            =   5520
      TabIndex        =   28
      Top             =   1680
      Value           =   1  'Checked
      Width           =   1830
   End
   Begin VB.TextBox TekstInfo 
      Height          =   315
      Index           =   5
      Left            =   2040
      TabIndex        =   13
      Top             =   1260
      Width           =   435
   End
   Begin VB.CommandButton DefEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TabPositie"
      Enabled         =   0   'False
      Height          =   405
      Index           =   3
      Left            =   4110
      TabIndex        =   27
      Top             =   4470
      Width           =   1245
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   4
      Left            =   5400
      TabIndex        =   26
      Top             =   5280
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      Height          =   315
      Index           =   3
      Left            =   5400
      TabIndex        =   25
      Top             =   5280
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
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   24
      Top             =   5280
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
      TabIndex        =   22
      Top             =   4020
      Width           =   1245
   End
   Begin VB.CommandButton DefEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Titel"
      Enabled         =   0   'False
      Height          =   405
      Index           =   1
      Left            =   4110
      TabIndex        =   21
      Top             =   3570
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
      TabIndex        =   20
      Top             =   3120
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
      TabIndex        =   19
      Top             =   3120
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
      TabIndex        =   17
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox TekstInfo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   1530
      TabIndex        =   15
      Top             =   2460
      Width           =   5925
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
      Height          =   300
      Left            =   6150
      TabIndex        =   11
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Afdrukken 
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
      Left            =   6150
      TabIndex        =   10
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox TekstInfo 
      Height          =   315
      Index           =   1
      Left            =   3210
      TabIndex        =   9
      Top             =   930
      Width           =   1635
   End
   Begin VB.TextBox TekstInfo 
      Height          =   315
      Index           =   0
      Left            =   3210
      TabIndex        =   7
      Top             =   600
      Width           =   1635
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
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   900
      Width           =   2415
   End
   Begin VB.CommandButton Definitie 
      Caption         =   "&Nieuwe Definitie maken "
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
      Left            =   3000
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1650
      Width           =   1875
   End
   Begin VB.ComboBox RapportDefinitie 
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
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   3585
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
      Left            =   60
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label lblAantalInSelektie 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      Height          =   285
      Left            =   3720
      TabIndex        =   31
      Top             =   1260
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Lengte Sorteersleutel"
      Height          =   195
      Index           =   8
      Left            =   90
      TabIndex        =   12
      Top             =   1350
      Width           =   1890
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Rapportvelden"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   7
      Left            =   5430
      TabIndex        =   23
      Top             =   2880
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
      TabIndex        =   18
      Top             =   2910
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
      TabIndex        =   16
      Top             =   2550
      Width           =   1275
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   60
      X2              =   9630
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Tot"
      Height          =   195
      Index           =   4
      Left            =   2880
      TabIndex        =   8
      Top             =   990
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Van"
      Height          =   195
      Index           =   3
      Left            =   2820
      TabIndex        =   6
      Top             =   660
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Sortering"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   690
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Aktieve &rapportdefinitie"
      Height          =   195
      Index           =   1
      Left            =   2550
      TabIndex        =   2
      Top             =   30
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Tabel"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   0
      Top             =   30
      Width           =   495
   End
End
Attribute VB_Name = "SQLLijsten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim TabLijn             As Integer
Dim FlKeuze             As Integer
Dim IndexKeuze          As Integer

Dim aa                  As String
Dim TLijnen             As Integer

Dim HoofdTitel          As String
Dim OnderTitel          As String

Dim rsLijst             As ADODB.Recordset

Sub SQLRefresh()

On Error Resume Next

Set rsLijst = New ADODB.Recordset
On Error Resume Next
Err = 0
rsLijst.CursorLocation = adUseClient
SnelHelpPrint sqlSELECTLijst, BL_LOGGING

rsLijst.Open sqlSELECTLijst, adntDB, adOpenForwardOnly, adLockReadOnly
If Err Then
    MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    Exit Sub
Else
    lblAantalInSelektie = Format(rsLijst.RecordCount, "#,##0")
End If

End Sub


Function sqlSELECTLijst() As String
Dim IndexeerVeld As String

IndexeerVeld = Trim$(JETTABLEUSE_INDEX(Val(Mid(Tabel.Text, 1, 1)), Val(Mid(Me.Sortering, 1, 2))))

Msg = "SELECT"
For COUNT_TO = 0 To Me.RapportVelden.ListCount - 1
    Msg = Msg + " " + Trim$(Mid(Me.RapportVelden.List(COUNT_TO), 5, 5))
    Msg = Msg + " AS [" + Trim$(Mid(RapportVelden.List(COUNT_TO), 22)) + "]"
    If COUNT_TO < Me.RapportVelden.ListCount - 1 Then
        Msg = Msg + ","
    End If
Next
Msg = Msg + " FROM " + Mid(Me.Tabel, 4)
Msg = Msg + " WHERE " + IndexeerVeld + " >= '" + Trim$(Me.TekstInfo(0)) + "'"
Msg = Msg + " AND " + IndexeerVeld + " <= '" + Trim$(Me.TekstInfo(1)) + "'"
Msg = Msg + " ORDER BY " + IndexeerVeld

EindeSelect:
sqlSELECTLijst = Msg

End Function

Private Sub Afdrukken_Click()
Dim BeginSleutel    As String
Dim EindSleutel     As String
Dim bModDummy       As Integer
Dim Lijn            As Integer
Dim DCBedrag        As Long
Dim XX              As Integer

aa = ""
TLijnen = 0
Lijn = 0
BeginSleutel = TekstInfo(0).Text
EindSleutel = TekstInfo(1).Text

HoofdTitel = Mid(Mim.Caption, InStr(Mim.Caption, "["))
OnderTitel = TekstInfo(2).Text
    
SQLRefresh
If rsLijst.RecordCount = 0 Then
    Exit Sub
Else
    rsLijst.MoveFirst
    Screen.MousePointer = vbHourglass
    PAGE_COUNTER = 0
    If chkAfdrukInVenster Then
    Else
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
    rsLijst.MoveNext
    If rsLijst.EOF Then
        Exit Do
    Else
        GoSub PrintInfo
    End If
Loop
PrintEinde
Printer.EndDoc
Screen.MousePointer = vbNormal
Sluiten.SetFocus
Exit Sub

PrintInfo:
Lijn = Lijn + 1
PrintVelden
Return

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
XDoEvents = DoEvents

End Sub

Private Sub cmdSQLOvername_Click()

frmADOSQLOpzoeker.Hide
frmADOSQLOpzoeker.txtSQL = rsLijst.Source
frmADOSQLOpzoeker.Show
frmADOSQLOpzoeker.txtSQL.SetFocus
SendKeys "{ENTER}"

End Sub

Private Sub cmdToonSQL_Click()

MsgBox sqlSELECTLijst

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

If Afdrukken.Enabled = True Then
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
    Afdrukken.Enabled = False
    TabelVelden.Enabled = True
    RapportVelden.Enabled = True
    TekstInfo(2).Enabled = True
    Tabel.Enabled = False
    Sortering.Enabled = False
    RapportDefinitie.Enabled = False
    DefEdit(0).Enabled = True
    WindowState = 2
    TekstInfo(2).SetFocus
Else
    Afdrukken.Enabled = True
    TabelVelden.Enabled = False
    RapportVelden.Enabled = False
    TekstInfo(2).Enabled = False
    Tabel.Enabled = True
    Sortering.Enabled = True
    RapportDefinitie.Enabled = True
    For T = 0 To 3
        DefEdit(T).Enabled = False
    Next
    WindowState = 0
    Dim FlTemp As Integer
    FlTemp = FreeFile
    VolgNummer% = RapportDefinitie.ListCount
    Open PROGRAM_LOCATION + "Def\" + Format(Val(Left(Tabel.Text, 1)), "000") + Format(VolgNummer%, "00") + ".PRD" For Output As FlTemp
    Print #FlTemp, (TekstInfo(2).Text)
    For T = 0 To RapportVelden.ListCount - 1
        Print #FlTemp, (RapportVelden.List(T))
    Next
    Close FlTemp
    RapportDefinitie.AddItem Format(VolgNummer%, "00") + ": " + TekstInfo(2).Text
    RapportDefinitie.ListIndex = RapportDefinitie.ListCount - 1
End If

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If Printer.Orientation = vbPRORLandscape Then
    chkAfdrukLiggend.Value = 1
Else
    chkAfdrukLiggend.Value = 0
End If
chkAfdrukLiggend_Click

Tabel.AddItem "1: Klanten"
Tabel.AddItem "2: Leveranciers"
Tabel.AddItem "3: Rekeningen"
Tabel.AddItem "4: Produkten"
Tabel.AddItem "5: Polissen"

Formattering.AddItem "T: Tekst zonder enige bewerking"
Formattering.AddItem "D: Van SorteerDatum naar DD/MM/EEJJ"
Formattering.AddItem "0: Bedrag met masker ########0"
Formattering.AddItem "1: Bedrag met masker ###0"
Formattering.AddItem "2: Bedrag met masker ######0.00"
Formattering.AddItem "3: Bedrag met masker ##0.00000000"
Formattering.AddItem "4: Bedrag met masker #######0.00"
Formattering.AddItem "5: Bedrag met masker ##0"
Formattering.AddItem "6: Bedrag met masker #0"
Formattering.AddItem "7: Bedrag met masker #####0.0"
Formattering.AddItem "Z: Rekenformule */+-() via Titel!"
Tabel.ListIndex = 0

'ADO TEST
'Dim cnn As adodb.Connection
'Dim rs As adodb.Recordset

'Set cnn = New adodb.Connection
'Set rs = New adodb.Recordset

'cnn.Open adojetprovider & "Data Source=C:\marnt\data\001\marnt.mdv;"
'rs.Open "SELECT * FROM Klanten", cnn, adOpenStatic, adLockOptimistic

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

Private Sub PrintEinde()

On Local Error Resume Next

If chkAfdrukInVenster Then
    Xlog.Caption = TekstInfo(2).Text
    Xlog!Kopie(1).Checked = True
    Xlog.X.Row = 1
    Xlog.X.Col = 0
    Xlog.WindowState = 2
    Screen.MousePointer = vbNormal
    Xlog.WijzigenLijn.Visible = False
    Xlog.Afsluiten.Enabled = False
    Xlog.Afsluiten.TabStop = False
    Xlog.cbAfbeelding.Visible = False
    XLogKey = ""
    Xlog.SSTab1.TabVisible(1) = False
    Xlog.Show 1
    Xlog.WindowState = 0
    Unload Xlog
End If

End Sub

Private Sub PrintTitel()
Dim T As Integer
Dim TabNu As Integer
Dim TabNext As Integer

If chkAfdrukInVenster = 1 And PAGE_COUNTER = 0 Then
    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = HoofdTitel
    Xlog.X.Cols = RapportVelden.ListCount
    Xlog.X.Row = 0
    For T = 0 To RapportVelden.ListCount - 1
        RapportVelden.ListIndex = T
        REPORT_TAB(T) = Val(Mid(RapportVelden.Text, 12, 3))
        Xlog.X.Col = T
        Xlog.X.Text = RTrim$(Mid(RapportVelden.Text, 22))
        Xlog.X.ColWidth(T) = 1860
    Next
    Me.Show
    Exit Sub
End If
    
    If usrLicentieInfo <> "" Then
        Printer.CurrentX = 50
        Printer.CurrentY = 50
        Printer.Print usrLicentieInfo;
    End If
    PAGE_COUNTER = PAGE_COUNTER + 1
    Printer.CurrentY = 400
    Printer.Print Tab(1); HoofdTitel; Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########");

    Printer.Print Tab(108); "Datum  : "; MIM_GLOBAL_DATE; vbCrLf; vbCrLf;
    Printer.Print Tab(1); OnderTitel;

Printer.Print vbCrLf; FULL_LINE

    For T = 0 To RapportVelden.ListCount - 1
        RapportVelden.ListIndex = T
        TabNu = Val(Mid(RapportVelden.Text, 12, 3))
        Printer.Print Tab(TabNu); Mid(RapportVelden.Text, 22);
    Next

Printer.Print vbCrLf; FULL_LINE; vbCrLf; vbCrLf;

End Sub

Private Sub PrintVelden()
Dim T As Integer
Dim VeldInfo As String
Dim VeldInfoSh As String

aa = ""
For T = 0 To RapportVelden.ListCount - 1
    On Error Resume Next
    VeldInfo = rsLijst.Fields(Mid(RapportVelden.List(T), 22))
    If Err Then VeldInfo = ""
    VeldInfoSh = Mid(RapportVelden.List(T), 20, 1)
    Select Case VeldInfoSh
        Case "T"
            VeldInfo = RTrim$(Left(VeldInfo, Val(Mid(RapportVelden.List(T), 16, 3))))
        Case "D"
             VeldInfo = DATE_TEXT((VeldInfo))
        Case "0" To "6"
            VeldInfo = Dec$(Val(VeldInfo), MASK_SY(Val(VeldInfoSh)))
        Case Else
            MsgBox Mid(RapportVelden.List(T), 20, 1) + " werd nog niet voorzien..."
    End Select
    If T = 0 Then
        SnelHelpPrint VeldInfo, BL_LOGGING
    End If
    If chkAfdrukInVenster Then
        aa = aa + VeldInfo
        If T <> RapportVelden.ListCount - 1 Then
            aa = aa + vbTab
        End If
    Else
        TabNu = Val(Mid(RapportVelden.List(T), 12, 3))
        Printer.Print Tab(TabNu); VeldInfo;
        If T = RapportVelden.ListCount - 1 Then
            Printer.Print vbCrLf;
        End If
    End If
Next
If chkAfdrukInVenster Then
    Xlog.X.AddItem aa, Xlog.X.Rows - 1
ElseIf Printer.CurrentY >= Printer.Height - 1200 Then
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If

End Sub


Private Sub RapportDefinitie_Click()

On Local Error Resume Next

    Dim FlTemp As Integer
    FlTemp = FreeFile
    
    If Dir(PROGRAM_LOCATION + "Def\" + Format(Val(Left(Tabel.Text, 1)), "000") + Left(RapportDefinitie.Text, 2) + ".PRD") = "" Then
        MsgBox Format(Val(Left(Tabel.Text, 1)), "000") + Left(RapportDefinitie.Text, 2) + ".PRD" + " bestaat niet meer..."
        Exit Sub
    End If
    
    Open PROGRAM_LOCATION + "Def\" + Format(Val(Left(Tabel.Text, 1)), "000") + Left(RapportDefinitie.Text, 2) + ".PRD" For Input As FlTemp
    Line Input #FlTemp, RapportTekst$
    TekstInfo(2).Text = RapportTekst$
    RapportVelden.Clear
    Do While Not EOF(FlTemp)
        Line Input #FlTemp, RapportTekst$
        RapportVelden.AddItem RapportTekst$
    Loop
    Close FlTemp

SQLRefresh

End Sub

Private Sub RapportDefinitie_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 46
        Msg = "Rapportdefinitie " + vbCrLf + vbCrLf + RapportDefinitie.Text + vbCrLf + vbCrLf + "verwijderen.  Bent U zeker ?"
        KtrlBox = MsgBox(Msg, 292)
        If KtrlBox = 6 Then
            Kill PROGRAM_LOCATION + "Def\" + Format(Val(Left(Tabel.Text, 1)), "000") + Left(RapportDefinitie.Text, 2) + ".PRD"
            RapportDefinitie.RemoveItem RapportDefinitie.ListIndex
            
        End If
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

Private Sub SorTering_Click()

IndexKeuze = Val(Left(Sortering.Text, 2))
TekstInfo(5).Text = Str$(FLINDEX_LEN(FlKeuze, IndexKeuze))
TekstInfo(0).Text = "0"
TekstInfo(1).Text = String$(Val(TekstInfo(5).Text), "z")

End Sub


Private Sub Tabel_Click()
Dim T As Integer
Dim FlL As Integer

Sortering.Clear
FlKeuze = Val(Left(Tabel.Text, 1))
For T = 0 To FL_NUMBEROFINDEXEN(FlKeuze)
    Sortering.AddItem Format(T, "00") + ":" + FLINDEX_CAPTION(FlKeuze, T)
Next
Sortering.ListIndex = 0
RapportDefinitie.Clear
X$ = Dir$(PROGRAM_LOCATION + "Def\" + Format(Val(Left(Tabel.Text, 1)), "000") + "??.PRD")
If X$ = "" Then
    Exit Sub
Else
    FlL = FreeFile
    Do While X$ <> ""
        Open PROGRAM_LOCATION + "Def\" + X$ For Input As FlL
            Line Input #FlL, TekstLijn$
        Close FlL
        RapportDefinitie.AddItem Mid(X$, 4, 2) + ": " + TekstLijn$
        X$ = Dir$
    Loop
End If
If RapportDefinitie.ListCount <> 0 Then
    RapportDefinitie.ListIndex = 0
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
SQLRefresh

End Sub

