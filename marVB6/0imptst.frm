VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form DatabaseTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Databases Beheer"
   ClientHeight    =   5205
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5205
   ScaleWidth      =   9420
   Begin VB.CheckBox ChkPrimair 
      Caption         =   "Primair"
      Height          =   225
      Left            =   7800
      TabIndex        =   15
      Top             =   3630
      Width           =   1485
   End
   Begin VB.CommandButton CmdVerwijderIndex 
      Caption         =   "VerwijderIndex"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton CmdMaakIndex 
      Caption         =   "MaakIndex"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   4320
      Width           =   1695
   End
   Begin VB.ComboBox CmbLijstTeZoeken 
      Enabled         =   0   'False
      Height          =   315
      Left            =   2850
      TabIndex        =   12
      Text            =   "CmbLijstTeZoeken"
      Top             =   1650
      Width           =   4725
   End
   Begin VB.CommandButton CmdZoek 
      Caption         =   "Zoek"
      Enabled         =   0   'False
      Height          =   315
      Left            =   1830
      TabIndex        =   11
      Top             =   1650
      Width           =   975
   End
   Begin VB.CheckBox ChkNoodzakelijk 
      Caption         =   "Noodzakelijk"
      Height          =   255
      Left            =   7800
      TabIndex        =   10
      Top             =   3300
      Width           =   1455
   End
   Begin VB.CheckBox ChkUniek 
      Caption         =   "Uniek"
      Height          =   255
      Left            =   7800
      TabIndex        =   9
      Top             =   2970
      Width           =   1455
   End
   Begin VB.CommandButton CmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   7680
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2340
      Width           =   1725
   End
   Begin VB.ListBox LstIndexen 
      Height          =   1230
      Left            =   7620
      TabIndex        =   7
      Top             =   360
      Width           =   1785
   End
   Begin VB.ListBox LstTabellen 
      Height          =   1230
      Left            =   4560
      TabIndex        =   3
      Top             =   360
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   990
      Left            =   60
      TabIndex        =   2
      Top             =   360
      Width           =   4365
   End
   Begin VB.CommandButton CmdDatabaseOpenen 
      Caption         =   "Drive en Path van database koppelen"
      Height          =   735
      Left            =   7680
      TabIndex        =   1
      Top             =   1620
      Width           =   1725
   End
   Begin VB.ComboBox CmbDatabasetypes 
      Height          =   315
      ItemData        =   "0imptst.frx":0000
      Left            =   60
      List            =   "0imptst.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   30
      Width           =   4305
   End
   Begin MSFlexGridLib.MSFlexGrid grdVelden 
      Height          =   2505
      Left            =   30
      TabIndex        =   16
      Top             =   1980
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   4419
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
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Beschikbare Indexen"
      Height          =   255
      Index           =   2
      Left            =   7650
      TabIndex        =   6
      Top             =   60
      Width           =   1635
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Beschikbare Velden"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   1680
      Width           =   1635
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Beschikbare Tabellen"
      Height          =   255
      Index           =   0
      Left            =   4590
      TabIndex        =   4
      Top             =   60
      Width           =   1635
   End
End
Attribute VB_Name = "DatabaseTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim NTIPDbas As DAO.Database
Dim NTIPTable As DAO.Recordset
Dim NTIPWorkspace As DAO.Workspace


Private Sub CmbLijstTeZoeken_Click()

Set NTIPTable = NTIPDbas.OpenRecordset(LstTabellen.Text, dbOpenTable)
NTIPTable.Index = LstIndexen.List(0)
NTIPTable.Seek "=", Mid(CmbLijstTeZoeken.Text, InStr(CmbLijstTeZoeken.Text, vbTab) + 1)
If NTIPTable.NoMatch Then
    grdVelden.Col = 3
    For TelTot = 0 To NTIPTable.Fields.Count - 1
        grdVelden.Row = TelTot + 1
        grdVelden.Text = ""
    Next
Else
    grdVelden.Col = 3
    On Local Error Resume Next
    For TelTot = 0 To NTIPTable.Fields.Count - 1
        grdVelden.Row = TelTot + 1
        grdVelden.Text = NTIPTable.Fields(TelTot)
    Next
End If
NTIPTable.Close

End Sub

Private Sub CmdDatabaseOpenen_Click()
        
LstTabellen.Clear
LstIndexen.Clear
grdVelden.Rows = 2
With grdVelden
    .Row = 1
    .Col = 0
    .Text = ""
    .Col = 1
    .Text = ""
    .Col = 2
    .Text = ""
End With

On Local Error Resume Next
NTIPDbas.Close

On Local Error GoTo CancelError
If Right(Dir1.path, 1) = "\" Then
    Mim.Teken.filename = Dir1.path + "*.*"
Else
    Mim.Teken.filename = Dir1.path + "\*.*"
End If
Mim.Teken.Filter = "Ms-Jet MDB (*.mdb,*,mdv)|*.mdb;*.mdv|dBase III, IV, 5.0 (*.dbf)|*.dbf|FoxPro 2.0, 2.5, 2.6 (*.dbf)|*.dbf|Paradox 3.x, 4.x, 5.x (*.db?)|*.db?"
Mim.Teken.CancelError = True
Mim.Teken.ShowOpen

On Local Error GoTo 0
Dim Pathnaam As String
Dim PathNaam2 As String
Dim ConnectString As String

Pathnaam = Mim.Teken.filename + " "
If Dir$(Pathnaam) = "" Then Exit Sub
Do While Right(Pathnaam, 1) <> "\"
    If Pathnaam = "" Then
        MsgBox "Stop"
    Else
        Pathnaam = Left(Pathnaam, Len(Pathnaam) - 1)
    End If
Loop
Pathnaam = Left(Pathnaam, Len(Pathnaam) - 1)
Dir1.path = Pathnaam
Dir1.Refresh
If UCase$(Right(Mim.Teken.filename, 3)) = "MDB" Or UCase$(Right(Mim.Teken.filename, 3)) = "MDV" Then
    ConnectString = ""
    PathNaam2 = Mim.Teken.filename
Else
    PathNaam2 = Pathnaam
    ConnectString = CmbDatabasetypes.Text + ";"
End If

Set NTIPWorkspace = Workspaces(0)

On Local Error GoTo OpenError
Set NTIPDbas = NTIPWorkspace.OpenDatabase(PathNaam2, False, False, ConnectString)

For T = 0 To NTIPDbas.TableDefs.Count - 1
    If NTIPDbas.TableDefs(T).Attributes <> 0 Then
    Else
        LstTabellen.AddItem NTIPDbas.TableDefs(T).Name
    End If
Next
Exit Sub

OpenError:
Select Case Err
    Case 3170
        MsgBox Error + vbCrLf + vbCrLf + "Kontroleer type en/of versie van de database a.u.b. !"
    Case Else
        MsgBox Error
End Select

CancelError:
LstTabellen.Clear

End Sub


Private Sub CmdMaakIndex_Click()
Dim IndexNaam As String

grdVelden.Col = 0
Msg = "Index : " + grdVelden.Text + vbCrLf + vbCrLf
If ChkNoodzakelijk.Value Then
    Msg = Msg + "Elk " + grdVelden.Text + " noodzakelijk aanwezig." + vbCrLf
Else
    Msg = Msg + "Elk " + grdVelden.Text + " hoeft NIET noodzakelijk aanwezig te zijn." + vbCrLf
End If
If ChkUniek.Value Then
    Msg = Msg + "Elk " + grdVelden.Text + " uniek." + vbCrLf
Else
    Msg = Msg + "Elk " + grdVelden.Text + " hoeft NIET uniek te zijn." + vbCrLf
End If
If ChkPrimair.Value Then
    Msg = Msg + grdVelden.Text + " is PRIMAIR." + vbCrLf
Else
    Msg = Msg + grdVelden.Text + " is NIET PRIMAIRE index." + vbCrLf
End If
Msg = Msg + vbCrLf + "Bent U zeker ?"
KtrlBox = MsgBox(Msg, vbYesNo)
If KtrlBox = vbYes Then
    Dim MyIndex As DAO.Index
    Dim MyField As DAO.Field
    Dim NieuwTable As DAO.TableDef
    Set NieuwTable = NTIPDbas.TableDefs(RTrim$(LstTabellen.Text))
    ' Create new Index object.
        IndexNaam = InputBox("Geef een naam aan de index", , RTrim$(grdVelden.Text))
        If IndexNaam = "" Then
            IndexNaam = RTrim$(grdVelden.Text)
        End If
        Set MyIndex = NieuwTable.CreateIndex(IndexNaam)
        Set MyField = MyIndex.CreateField(RTrim$(grdVelden.Text))
    MyIndex.Primary = ChkPrimair.Value
    MyIndex.Required = ChkNoodzakelijk.Value
    MyIndex.Unique = ChkUniek.Value
    On Local Error Resume Next
    Screen.MousePointer = vbHourglass
    MyIndex.Fields.Append MyField
    ' Save Index definition by appending it to Indexes collection.
    NieuwTable.Indexes.Append MyIndex
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox Error
    Else
        LstTabellen_Click
    End If
End If

End Sub

Private Sub cmdSluiten_Click()

Unload Me

End Sub

Private Sub CmdVerwijderIndex_Click()

If LstIndexen.ListIndex < 0 Then Exit Sub
Msg = "Index : " + LstIndexen.Text + " verwijderen." + vbCrLf + vbCrLf
Msg = Msg + "Bent U zeker ?"
KtrlBox = MsgBox(Msg, vbYesNo + vbDefaultButton2)
If KtrlBox = vbYes Then
    NTIPDbas.TableDefs(LstTabellen.Text).Indexes.Delete LstIndexen.Text
    LstTabellen_Click
End If

End Sub

Private Sub CmdZoek_Click()
Dim ZoekTekst As String

If CmbLijstTeZoeken.Text = "" Then
    MsgBox "Minstens 1 teken a.u.b. !"
Else
    Screen.MousePointer = vbHourglass
    ZoekTekst = CmbLijstTeZoeken.Text
    CmbLijstTeZoeken.Clear
    Set NTIPTable = NTIPDbas.OpenRecordset(LstTabellen.Text, dbOpenTable)
    NTIPTable.Index = LstIndexen.Text
    NTIPTable.Seek ">=", ZoekTekst
    If NTIPTable.NoMatch Then
    Else
        On Local Error Resume Next
        Do
            If UCase(ZoekTekst) <> UCase(Left(NTIPTable.Fields(LstIndexen.Text), Len(ZoekTekst))) Then
                Exit Do
            Else
                On Local Error Resume Next
                CmbLijstTeZoeken.AddItem NTIPTable.Fields(LstIndexen.Text) + vbTab + NTIPTable.Fields(0)
                On Local Error GoTo 0
                NTIPTable.MoveNext
            End If
        Loop
    End If
    NTIPTable.Close
    Screen.MousePointer = vbNormal
    If CmbLijstTeZoeken.ListCount Then
        CmbLijstTeZoeken.ListIndex = 0
    End If
End If

End Sub

Private Sub Form_DblClick()

Dim FlTemp          As Integer
Dim TT              As Integer
Dim T               As Integer

Dim RecordString    As String

Top = 0
Left = 0
'Dim ntDBase As Database
Dim NTTable As DAO.Recordset

On Local Error GoTo CancelError
If Right(Dir1.path, 1) = "\" Then
    Mim.Teken.filename = Dir1.path + "*.*"
Else
    Mim.Teken.filename = Dir1.path + "\*.*"
End If
Mim.Teken.Filter = "Data-Tekstbestanden (*.txt)|*.txt"
Mim.Teken.CancelError = True
Mim.Teken.ShowOpen

On Local Error GoTo 0
GoTo HierBegintHetDan

CancelError:
Screen.MousePointer = vbNormal
Exit Sub

HierBegintHetDan:
ReDim Velden(100) As String

DatabaseTest.Enabled = False
Screen.MousePointer = vbHourglass
FlTemp = FreeFile
Open Mim.Teken.filename For Input As FlTemp
    Line Input #FlTemp, RecordString
    GoSub TitelKop
    BeginTrans
    Do While Not EOF(FlTemp)
        XDoEvents = DoEvents
        Line Input #FlTemp, RecordString
        GoSub VoegErBij
    Loop
Close FlTemp
CommitTrans
NTTable.Close
'ntDBase.Close
DatabaseTest.Enabled = True
Screen.MousePointer = vbNormal
MsgBox "Invoeging gegevensverzameling : " + Mid(Velden(0), 1, InStr(Velden(0), ".") - 1) + " met succes beëindigd."
Exit Sub

TitelKop:
T = 0
Do While RecordString <> ""
    If InStr(RecordString, vbTab) <> 0 Then
        Velden(T) = Mid(RecordString, 1, InStr(RecordString, vbTab) - 1)
        RecordString = Mid(RecordString, InStr(RecordString, vbTab) + 1)
        T = T + 1
    Else
        Velden(T) = RecordString
        RecordString = ""
    End If
Loop
GoSub DbKontrole
'Set ntDBase = ntruimte.OpenDatabase(ProgrammaLokatie + "default\Default2018.mdb")
Set NTTable = ntDB.OpenRecordset(Mid(Velden(0), 1, InStr(Velden(0), ".") - 1), dbOpenTable)
Return

VoegErBij:
NTTable.AddNew
T = 0
Do While RecordString <> ""
    If InStr(RecordString, vbTab) <> 0 Then
        VeldZelf = Mid(RecordString, 1, InStr(RecordString, vbTab) - 1)
        If VeldZelf = "" Then
        Else
            NTTable.Fields(Mid(Velden(T), InStr(Velden(T), ".") + 1)) = VeldZelf
        End If
        RecordString = Mid(RecordString, InStr(RecordString, vbTab) + 1)
        T = T + 1
    Else
        VeldZelf = RecordString
        If VeldZelf <> "" Then
            NTTable.Fields(Mid(Velden(T), InStr(Velden(T), ".") + 1)) = VeldZelf
        End If
        RecordString = ""
    End If
Loop
SnelHelpPrint NTTable.Fields(0), blLogging
NTTable.Update
Return

DbKontrole:
Set ntDBase = NTRuimte.OpenDatabase(ProgrammaLokatie + "Default2018.mdb")
For TT = 0 To ntDB.TableDefs.Count - 1
    If ntDBase.TableDefs(TT).Name = Mid(Velden(0), 1, InStr(Velden(0), ".") - 1) Then
        MsgBox "Bestaande gegevensverzameling : " + Mid(Velden(0), 1, InStr(Velden(0), ".") - 1) + " wordt hierna eerst verwijderd."
        ntDBase.TableDefs.Delete Mid(Velden(0), 1, InStr(Velden(0), ".") - 1)
        Exit For
    End If
Next
KtrlBox = MsgBox("Nieuwe gegevensverzameling : " + Mid(Velden(0), 1, InStr(Velden(0), ".") - 1) + " invoegen.  Bent U zeker ?", vbYesNo + vbQuestion + vbDefaultButton2)
If KtrlBox = vbNo Then
    Screen.MousePointer = vbNormal
    DatabaseTest.Enabled = True
    Exit Sub
End If

Dim TD() As New DAO.TableDef, fld() As New DAO.Field
Dim Idx() As New DAO.Index, i As Integer
Dim FlHier As Integer
ReDim TD(1)
ReDim fld(0 To T + 1), Idx(0 To 2)
    TD(0).Name = Mid(Velden(0), 1, InStr(Velden(0), ".") - 1)    ' Set the table name.
    ' Create Fields.
    
    For i = 0 To T  ' Set properties for fields.
        fld(i).Name = Mid(Velden(i), InStr(Velden(i), ".") + 1)  'Choose(I, "ANum", "Addr1", "Addr2", "CityState", "ZIP")
        fld(i).Type = 10                'Choose(I, DB_LONG, DB_TEXT, DB_TEXT, DB_TEXT, DB_TEXT)
        If i = 2 Then
            fld(i).Size = 100
        Else
            fld(i).Size = 40
        End If
        TD(0).Fields.Append fld(i)
    Next i
    ntDB.TableDefs.Append TD(0)
'ntDBase.Close
Return

End Sub

Private Sub Form_Load()

'DatabaseTest.WindowState = 2

grdVelden.Cols = 4
grdVelden.Row = 0
grdVelden.Col = 0
With grdVelden
    .Col = 0
    .Text = "Naam"
    .ColWidth(0) = 1680
    
    .Col = 1
    .Text = "Lengte"
    .ColAlignment(1) = 1
    .ColWidth(1) = 525
        
    .Col = 2
    .Text = "Type"
    .ColWidth(2) = 1260
        
    .Col = 3
    .Text = "VeldInfo"
    .ColWidth(3) = 3675
End With
With CmbDatabasetypes
    .AddItem "Jet MDB"
    .AddItem "dBASE III"
    .AddItem "dBASE IV"
    .AddItem "dBASE 5.0"
    .AddItem "FoxPro 2.0"
    .AddItem "FoxPro 2.5"
    .AddItem "FoxPro 2.6"
    .AddItem "Paradox 3.X"
    .AddItem "Paradox 4.X"
    .AddItem "Paradox 5.X"
    .ListIndex = 0
End With

End Sub


Private Sub Form_Unload(Cancel As Integer)

On Local Error Resume Next

NTIPDbas.Close

End Sub


Private Sub LstIndexen_Click()

On Local Error Resume Next
ChkPrimair.Value = NTIPDbas.TableDefs(LstTabellen.Text).Indexes(Val(LstIndexen.ListIndex)).Primary
ChkNoodzakelijk.Value = NTIPDbas.TableDefs(LstTabellen.Text).Indexes(LstIndexen.Text).Required
ChkUniek.Value = NTIPDbas.TableDefs(LstTabellen.Text).Indexes(LstIndexen.Text).Unique
CmdVerwijderIndex.Enabled = True
CmdZoek.Enabled = True
CmbLijstTeZoeken.Enabled = True
CmbLijstTeZoeken.Clear

End Sub

Private Sub LstTabellen_Click()

Dim VeldType(1 To 13) As String

VeldType(1) = "01: Boolean (Ja/nee)"
VeldType(2) = "02: Byte"
VeldType(3) = "03: Geheel getal"
VeldType(4) = "04: Lang geheel getal"
VeldType(5) = "05: Muntgetal"
VeldType(6) = "06: Enkel nauwkeurig"
VeldType(7) = "07: Dubbel nauwkeurig"
VeldType(8) = "08: Datum"
VeldType(10) = "10: Tekst"
VeldType(11) = "11: Lang binair"
VeldType(12) = "12: Memo"

grdVelden.Rows = 2
With grdVelden
    .Row = 1
    .Col = 0
    .Text = ""
    .Col = 1
    .Text = ""
    .Col = 2
    .Text = ""
End With
For T = 0 To NTIPDbas.TableDefs(LstTabellen.Text).Fields.Count - 1
    A = NTIPDbas.TableDefs(LstTabellen.Text).Fields(T).Name + vbTab
    If NTIPDbas.TableDefs(LstTabellen.Text).Fields(T).Type = 10 Then
        A = A + Str$(NTIPDbas.TableDefs(LstTabellen.Text).Fields(T).Size) + vbTab
    Else
        A = A + "--" + vbTab
    End If
    A = A + VeldType(NTIPDbas.TableDefs(LstTabellen.Text).Fields(T).Type)
    grdVelden.AddItem A, grdVelden.Rows - 1
Next
LstIndexen.Clear
On Local Error Resume Next
For T = 0 To NTIPDbas.TableDefs(LstTabellen.Text).Indexes.Count - 1
    LstIndexen.AddItem NTIPDbas.TableDefs(LstTabellen.Text).Indexes(T).Name
Next
CmdMaakIndex.Enabled = True
CmdVerwijderIndex.Enabled = False
CmdZoek.Enabled = False
CmbLijstTeZoeken.Enabled = False
If LstIndexen.ListCount Then
    LstIndexen.ListIndex = 0
End If

End Sub



