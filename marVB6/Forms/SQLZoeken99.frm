VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form SqlSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ANSI-92 SQL GeSELECTeerd zoeken"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7740
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar stbSnelHelp 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   10
      Top             =   4080
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Sortering 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   480
      Width           =   4455
   End
   Begin RichTextLib.RichTextBox rtbSQLTekst 
      Height          =   855
      Left            =   0
      TabIndex        =   8
      TabStop         =   0   'False
      ToolTipText     =   "Verfijn de SQL instructie naar believen !"
      Top             =   3000
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   1508
      _Version        =   393217
      BackColor       =   16777215
      TextRTF         =   $"SQLZoeken99.frx":0000
   End
   Begin VB.CommandButton cmdZoeken 
      Caption         =   "Zoeken"
      Default         =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   6480
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdBewaar 
      Caption         =   "Bewaren"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1215
   End
   Begin VB.ComboBox cmbExternedatabase 
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   480
      Width           =   4455
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgLijst 
      Height          =   2055
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3625
      _Version        =   393216
      BackColor       =   -2147483624
      ForeColor       =   0
      FixedCols       =   0
      FocusRect       =   2
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   0
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CheckBox chkExterneDatabase 
      Caption         =   "E&xterne database"
      Height          =   195
      Left            =   4560
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtTeZoeken 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label lblTekst 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   9
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label lblTekst 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " &Zoek zoals"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "SqlSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim IndexNR As Integer
Dim grdColWidth(20) As Integer

Private datPrimaryRS As ADODB.Recordset

Function SQLVernieuwTekst(ComboTekst As String)
Dim SorteerIndex As String
Dim SorteerOrde As String
Dim Sleuteltje As String
Dim TelOrde As Integer

On Error GoTo ErrorSQLVernieuwTekst

grdColWidth(0) = 0
'eerst kontrole voorkeurSQL
If chkExterneDatabase.Value = 0 Then
    Sleuteltje = "marSQL" + Format(SharedFl, "00") + Left(ComboTekst, InStr(ComboTekst, ";") - 1)
Else
    Sleuteltje = "marEDB" + Format(SharedFl, "00") + Left(ComboTekst, InStr(ComboTekst, ";") - 1)
End If
TelOrde = 0
Do
    COUNT_TO = InStr(TelOrde + 1, ComboTekst, ";") - 1
    If COUNT_TO < 0 Then
        Exit Do
    Else
        If TelOrde = 0 Then
            SorteerIndex = Mid(ComboTekst, COUNT_TO - 3, 4)
            SorteerOrde = Mid(ComboTekst, COUNT_TO - 3, 4)
            If Mid(ComboTekst, COUNT_TO - 4, 1) = "+" Then
                SorteerOrde = SorteerOrde + " ASC"
            Else
                SorteerOrde = SorteerOrde + " DESC"
            End If
        Else
            SorteerIndex = SorteerIndex + "+" + Mid(ComboTekst, COUNT_TO - 3, 4)
            SorteerOrde = SorteerOrde + ", " + Mid(ComboTekst, COUNT_TO - 3, 4)
            If Mid(ComboTekst, COUNT_TO - 4, 1) = "+" Then
                SorteerOrde = SorteerOrde + " ASC"
            Else
                SorteerOrde = SorteerOrde + " DESC"
            End If
        End If
    
    End If
    TelOrde = COUNT_TO + 1
Loop
   
bGet TABLE_VARIOUS, 1, "29" + Sleuteltje
If Ktrl Then
    If chkExterneDatabase.Value = 1 Then
        MsgBox "Onjuiste SQL zoekopdracht voor externe database", vbCritical
        Exit Function
    Else
        GoSub InitSQL
    End If
Else
    RecordToVeld TABLE_VARIOUS
    Msg = vBibTekst(TABLE_VARIOUS, "#v132 #")
    If InStr(UCase(Msg), "WHERE") Then
        Msg = Left(Msg, InStr(UCase(Msg), " WHERE ") - 1)
        Msg = Msg + " WHERE " + SorteerIndex + " Like " + Chr(34) + txtTeZoeken + Chr(34)
        Msg = Msg + " ORDER BY " + SorteerOrde
        rtbSQLTekst = Msg
        Msg = Mid(vBibTekst(TABLE_VARIOUS, "#v132 #"), InStr(vBibTekst(TABLE_VARIOUS, "#v132 #"), "[Colwidth]") + 10)
        If Msg = "" Then
            grdColWidth(0) = 0
        Else
            On Error Resume Next
            COUNT_TO = 0
            Do While Msg <> ""
                If InStr(Msg, vbTab) <> 0 Then
                    grdColWidth(COUNT_TO) = Val(Left(Msg, InStr(Msg, vbTab) - 1))
                    Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
                    COUNT_TO = COUNT_TO + 1
                Else
                    Exit Do
                End If
            Loop
            grdColWidth(COUNT_TO) = 0
        End If
    Else
        GoSub InitSQL
    End If
End If
Exit Function

InitSQL:
Msg = "SELECT"
Dim Delaatste As Boolean

Delaatste = False
'eerst eerste index verzekeren !
For COUNT_TO = 0 To Sortering.ListCount - 1
    If Trim(JETTABLEUSE_INDEX(SharedFl, 0)) = Mid(Sortering.List(COUNT_TO), 2, InStr(Sortering.List(COUNT_TO), ";") - 2) Then
        Msg = Msg + " " + Mid(Sortering.List(COUNT_TO), 2, InStr(Sortering.List(COUNT_TO), ";") - 2)
        Msg = Msg + " AS [" + Mid(Sortering.List(COUNT_TO), InStr(Sortering.List(COUNT_TO), ";") + 2) + "]"
        Msg = Msg + ","
        If COUNT_TO = Sortering.ListCount - 1 Then
            Delaatste = True
        End If
        Exit For
    End If
Next
If Msg = "SELECT" Then
    MsgBox "Hoofdindex " & Mid(Sortering.List(COUNT_TO), 2, InStr(Sortering.List(COUNT_TO), ";") - 2) & " bestaat niet (meer)", vbCritical
End If

'dan de rest bijvoegen
For COUNT_TO = 0 To Sortering.ListCount - 1
    If Trim(JETTABLEUSE_INDEX(SharedFl, 0)) = Mid(Sortering.List(COUNT_TO), 2, InStr(Sortering.List(COUNT_TO), ";") - 2) Then
    Else
        Msg = Msg + " " + Mid(Sortering.List(COUNT_TO), 2, InStr(Sortering.List(COUNT_TO), ";") - 2)
        Msg = Msg + " AS [" + Mid(Sortering.List(COUNT_TO), InStr(Sortering.List(COUNT_TO), ";") + 2) + "]"
        If Delaatste = True And Sortering.ListCount = 1 Then
        ElseIf Delaatste = True And COUNT_TO = Sortering.ListCount - 2 Then
        ElseIf COUNT_TO < Sortering.ListCount - 1 Then
            Msg = Msg + ","
        End If
    End If
Next
Msg = Msg + " FROM " + bstNaam(SharedFl)
Msg = Msg + " WHERE " + SorteerIndex + " Like " + Chr(34) + txtTeZoeken + Chr(34)
Msg = Msg + " ORDER BY " + SorteerOrde
rtbSQLTekst = Msg
Return

ErrorSQLVernieuwTekst:
MsgBox "Een fout tijdens opbouw van de SQL SELECT instructie."

End Function

Private Sub chkExterneDataBase_Click()

If chkExterneDatabase.Value = 1 Then
    Sortering.Visible = False
    cmbExternedatabase.Visible = True
    VulcmbExterneDatabase
Else
    cmbExternedatabase.Visible = False
    VulcmbSortering
    Sortering.Visible = True
End If

End Sub

Private Sub cmbExterneDatabase_Click()

cmdZoeken.Caption = "Zoeken"
SQLVernieuwTekst cmbExternedatabase.text
Schoon
'txtTeZoeken.SetFocus

End Sub

Private Sub CmdBewaar_Click()
Dim Sleuteltje As String

If chkExterneDatabase.Value = 0 Then
    Sleuteltje = "marSQL" + Format(SharedFl, "00") + Left(Sortering.text, InStr(Sortering.text, ";") - 1)
Else
    Sleuteltje = "marEDB" + Format(SharedFl, "00") + Left(cmbExternedatabase.text, InStr(cmbExternedatabase.text, ";") - 1)
End If

bGet TABLE_VARIOUS, 1, "29" + Sleuteltje
If Ktrl Then
    Msg = ""
    For COUNT_TO = 0 To mfgLijst.Cols - 1
        Msg = Msg + Format(mfgLijst.ColWidth(COUNT_TO)) + vbTab
    Next
    Msg = "[Colwidth]" + Msg
    Msg = rtbSQLTekst.text + Msg
    
    'Als nieuw bijvoegen
    TLB_RECORD(TABLE_VARIOUS) = ""
    vBib TABLE_VARIOUS, Msg, "v132"
    vBib TABLE_VARIOUS, Sleuteltje, "v250"
    vBib TABLE_VARIOUS, "29" + Sleuteltje, "v005"
    bInsert TABLE_VARIOUS, 1
ElseIf MsgBox("Bestaande definitie '" + Sleuteltje + "' overschrijven ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    RecordToVeld TABLE_VARIOUS
    Msg = ""
    For COUNT_TO = 0 To mfgLijst.Cols - 1
        Msg = Msg + Format(mfgLijst.ColWidth(COUNT_TO)) + vbTab
    Next
    Msg = "[Colwidth]" + Msg
    Msg = rtbSQLTekst.text + Msg
    vBib TABLE_VARIOUS, Msg, "v132"
    vBib TABLE_VARIOUS, Sleuteltje, "v250"
    vBib TABLE_VARIOUS, "29" + Sleuteltje, "v005"
    bUpdate TABLE_VARIOUS, 1
Else
    Exit Sub
End If

End Sub

Private Sub cmdSluiten_Click()

Ktrl = 99
Unload Me

End Sub


Private Sub Form_Load()

Dim T As Integer
Dim SorteringTel As Integer

SorteringTel = 0
Caption = Caption + ": " + bstNaam(SharedFl)
VulcmbSortering

If InStr(GridText, "@Beperk@") Then
    txtTeZoeken.text = Left(GridText, 2) + "%"
    cmdZoeken_Click
ElseIf GridText <> "" Then
    txtTeZoeken.text = GridText + "%"
    cmdZoeken_Click
Else
    txtTeZoeken.text = "%"
End If

End Sub


Private Sub mfgLijst_Click()

On Error Resume Next
datPrimaryRS.MoveFirst
datPrimaryRS.Move mfgLijst.Row - 1
txtTeZoeken.text = datPrimaryRS(Mid(Sortering, InStr(Sortering, ";") + 2))
If Err Then
    SnelHelpPrint Mid(Sortering, InStr(Sortering, ";") + 2) + " ontbreekt in SELECT !!", False
End If

End Sub

Private Sub mfgLijst_DblClick()

cmdZoeken_Click

End Sub

Private Sub mfgLijst_GotFocus()

cmdZoeken.Caption = "Ok"
mfgLijst_Click

End Sub


Private Sub mfgLijst_RowColChange()

If mfgLijst.Rows <> 2 Then
    mfgLijst_Click
End If

End Sub



Private Sub rtbSQLTekst_KeyPress(KeyAscii As Integer)

cmdBewaar.Enabled = True

End Sub

Private Sub SorTering_Click()

On Error Resume Next
cmdZoeken.Caption = "Zoeken"
SQLVernieuwTekst Sortering.text
txtTeZoeken = "%"
Schoon
'txtTeZoeken.SetFocus

End Sub

Private Sub Sortering_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case 38, 40
            
        Case Else
            KeyCode = 0
    End Select
    
End Sub

Private Sub Sortering_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub txtTeZoeken_Change()

If Len(txtTeZoeken) <= 1 And InStr(txtTeZoeken, "%") = 0 Then
    txtTeZoeken = txtTeZoeken + "%"
    txtTeZoeken.SelStart = Len(txtTeZoeken) - 1
End If

End Sub

Private Sub txtTeZoeken_GotFocus()

cmdZoeken.Caption = "Zoeken"
txtTeZoeken.SelLength = Len(txtTeZoeken)

End Sub

Private Sub txtTeZoeken_KeyDown(KeyCode As Integer, Shift As Integer)

On Error Resume Next

Select Case KeyCode
    Case 38 To 40
        mfgLijst.SetFocus
        mfgLijst.Col = IndexNR
        mfgLijst.Row = 1
End Select

End Sub

Private Sub cmdZoeken_Click()

Dim RecordIsNieuw As Boolean
Dim AlleRecordsInvoegen As Boolean
Dim RijInteVoegen As Long
Dim TellerHier As Long

If cmdZoeken.Caption = "Ok" Then
    If Sortering.Visible = True Then
        mfgLijst.Col = 0
        XLogKey = mfgLijst
        bGet SharedFl, 0, mfgLijst.text
        If Ktrl Then
            Beep
            txtTeZoeken.SetFocus
            Exit Sub
        Else
            RecordToVeld SharedFl
        End If
    Else
        KtrlBox = MsgBox("Info uit externe database toevoegen aan bedrijfsdatabase.  Bent U zeker?" + vbCrLf + vbCrLf + "Kies 'Ja' om enkel geselecteerde lijn in te voegen." + vbCrLf + "Kies 'Nee' om alle lijnen in te voegen.", vbQuestion + vbYesNoCancel + vbDefaultButton3, mfgLijst.TextMatrix(mfgLijst.Row, 0))
        If KtrlBox = vbYes Then
            RijInteVoegen = mfgLijst.Row
            GoSub VoegExterneLijnErBij
            chkExterneDatabase.Value = 0
            txtTeZoeken.text = ""
            txtTeZoeken.SetFocus
            Exit Sub
        ElseIf KtrlBox = vbNo Then
            KtrlBox = MsgBox("Alle lijnen invoegen.  Bent U zeker?", vbQuestion + vbYesNo + vbDefaultButton2)
            If KtrlBox = vbYes Then
                RijInteVoegen = -1
                GoSub VoegExterneLijnErBij
                chkExterneDatabase.Value = 0
                txtTeZoeken.text = ""
                txtTeZoeken.SetFocus
                Exit Sub
            End If
        Else
            Exit Sub
        End If
    End If
    Unload Me
Else
    If chkExterneDatabase.Value = 0 Then
        SQLVernieuwTekst Sortering.text
    Else
        SQLVernieuwTekst cmbExternedatabase.text
    End If
    Me.MousePointer = vbHourglass
    
    Dim sSQL As String
    sSQL = rtbSQLTekst.text
    
    ' Create a recordset using the provided collection
    Set datPrimaryRS = New ADODB.Recordset
    datPrimaryRS.CursorLocation = adUseClient
    On Error Resume Next
    datPrimaryRS.Open sSQL, adntDB 'Is toch optioneel: , adOpenForwardOnly, adLockReadOnly
    If Err.Number = -2147217904 Then
        Set mfgLijst.DataSource = Nothing
    ElseIf Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        Set mfgLijst.DataSource = Nothing
    Else
        Set mfgLijst.DataSource = datPrimaryRS
    End If
    lblTekst(1) = datPrimaryRS.RecordCount
       
    For COUNT_TO = 0 To mfgLijst.Cols - 1
        If grdColWidth(COUNT_TO) = 0 Then
            Exit For
        Else
            mfgLijst.ColWidth(COUNT_TO) = grdColWidth(COUNT_TO)
        End If
    Next
     
    Me.MousePointer = vbNormal
    If mfgLijst.Rows > 1 Then
        mfgLijst.Row = 1
        mfgLijst.Col = 0
        On Error Resume Next
        mfgLijst.SetFocus
    End If
End If
Exit Sub

VoegExterneLijnErBij:
If Trim(JETTABLEUSE_INDEX(SharedFl, 0)) <> Trim(mfgLijst.TextArray(0)) Then
    MsgBox "Eerste veldnaam komt niet overeen"
    Exit Sub
ElseIf RijInteVoegen = -1 Then
    Screen.MousePointer = vbHourglass
    For TellerHier = 1 To mfgLijst.Rows - 1
        bGet SharedFl, 0, mfgLijst.TextMatrix(TellerHier, 0)
        If Ktrl Then
            RecordIsNieuw = True
            KtrlBox = daoBlankoRecord(SharedFl)
        Else
            RecordIsNieuw = False
            stbSnelHelp.SimpleText = "Bestaat reeds"
            RecordToVeld SharedFl
        End If
        For T = 0 To mfgLijst.Cols - 1
            vBib SharedFl, mfgLijst.TextMatrix(TellerHier, T), mfgLijst.TextMatrix(0, T)
        Next
        If RecordIsNieuw Then
            bInsert SharedFl, 0
        Else
            bUpdate SharedFl, 0
        End If
    Next
    Screen.MousePointer = vbNormal
Else
    bGet SharedFl, 0, mfgLijst.TextMatrix(RijInteVoegen, 0)
    If Ktrl Then
        RecordIsNieuw = True
        KtrlBox = daoBlankoRecord(SharedFl)
    Else
        RecordIsNieuw = False
        stbSnelHelp.SimpleText = "Bestaat reeds"
        RecordToVeld SharedFl
    End If
    For T = 0 To mfgLijst.Cols - 1
        vBib SharedFl, mfgLijst.TextMatrix(RijInteVoegen, T), mfgLijst.TextMatrix(0, T)
    Next
    If RecordIsNieuw Then
        bInsert SharedFl, 0
    Else
        bUpdate SharedFl, 0
    End If
End If
Return

SQLZoekFout:
Me.MousePointer = vbNormal
MsgBox Err.Description

End Sub

Private Sub txtTeZoeken_KeyPress(KeyAscii As Integer)

cmdZoeken.Caption = "Zoeken"

End Sub

Private Sub VulcmbSortering()

GetAllIndexes bstNaam(SharedFl), Sortering
For T = 0 To Sortering.ListCount - 1
    If Mid(Sortering.List(T), InStr(Sortering.List(T), ";") + 2) = FLINDEX_CAPTION(SharedFl, aIndex) Then
        IndexNR = T
        Exit For
    End If
Next
If IndexNR Then
    Sortering.ListIndex = IndexNR
ElseIf Sortering.ListCount Then
    Sortering.ListIndex = 0
End If

End Sub

Private Sub VulcmbExterneDatabase()

Dim VeldString As String
Dim SleutelHier As String
    
cmbExternedatabase.Clear
SleutelHier = "marEDB" + Format(SharedFl, "00")
bGetOrGreater TABLE_VARIOUS, 1, "29" + SleutelHier
If Ktrl Or UCase(Mid(KEY_BUF(TABLE_VARIOUS), 3, 8)) <> UCase(SleutelHier) Then
    MsgBox "Er bestaan (nog) geen definities met voorvoegsel: " & SleutelHier, vbExclamation
    chkExterneDatabase.Value = 0
Else
    Do
        RecordToVeld TABLE_VARIOUS
        VeldString = Mid(KEY_BUF(TABLE_VARIOUS), 11, 5)
        cmbExternedatabase.AddItem VeldString + ";" + Mid(KEY_BUF(TABLE_VARIOUS), 3)
        bNext TABLE_VARIOUS
        If Ktrl Or UCase(Mid(KEY_BUF(TABLE_VARIOUS), 3, 8)) <> UCase(SleutelHier) Then
            Exit Do
        End If
    Loop
    cmbExternedatabase.ListIndex = 0
End If

End Sub

Sub Schoon()

    mfgLijst.Clear
    mfgLijst.Cols = 2
    mfgLijst.Rows = 2
    
End Sub
