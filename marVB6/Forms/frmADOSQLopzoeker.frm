VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmADOSQLOpzoeker 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ANSI-92 SQL Database Beheer"
   ClientHeight    =   5010
   ClientLeft      =   465
   ClientTop       =   3795
   ClientWidth     =   9135
   HelpContextID   =   11500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   9135
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Backup"
      Height          =   375
      Left            =   6720
      TabIndex        =   18
      Top             =   1860
      Width           =   675
   End
   Begin VB.CommandButton cmdNet1 
      Caption         =   "Net1"
      Height          =   375
      Left            =   6180
      TabIndex        =   17
      Top             =   1860
      Width           =   495
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msfSQL 
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   3836
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtPLUS 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1440
      TabIndex        =   16
      Top             =   4500
      Width           =   3015
   End
   Begin VB.TextBox txtWaarde 
      Enabled         =   0   'False
      Height          =   255
      Left            =   7560
      TabIndex        =   15
      Top             =   4560
      Width           =   1515
   End
   Begin VB.ComboBox cbOperatie 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmADOSQLopzoeker.frx":0000
      Left            =   5940
      List            =   "frmADOSQLopzoeker.frx":0019
      TabIndex        =   14
      Top             =   4500
      Width           =   1575
   End
   Begin VB.ComboBox cbVelden 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmADOSQLopzoeker.frx":0038
      Left            =   4500
      List            =   "frmADOSQLopzoeker.frx":003A
      TabIndex        =   13
      Top             =   4500
      Width           =   1395
   End
   Begin VB.ComboBox cbSQLBevel 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmADOSQLopzoeker.frx":003C
      Left            =   60
      List            =   "frmADOSQLopzoeker.frx":004F
      TabIndex        =   12
      Top             =   4500
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpenXML 
      Caption         =   "XML &Openen"
      Height          =   690
      Left            =   6180
      Picture         =   "frmADOSQLopzoeker.frx":007E
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "XML RecordSet ophalen"
      Top             =   1080
      Width           =   1200
   End
   Begin VB.CommandButton cmdSelectWegschrijven 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4320
      Picture         =   "frmADOSQLopzoeker.frx":06E8
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2310
      Width           =   375
   End
   Begin VB.ComboBox cmbSelect 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   2310
      Width           =   2835
   End
   Begin VB.TextBox txtSQL 
      Height          =   1755
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2640
      Width           =   7305
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "SQL &EXECUTE"
      Height          =   300
      Left            =   4740
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1350
   End
   Begin VB.CommandButton cmdKopij 
      Caption         =   "XML &Kopie"
      Height          =   690
      Left            =   6180
      Picture         =   "frmADOSQLopzoeker.frx":07EA
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "XML RecordSet maken of kopiëren naar klassiek plakbord"
      Top             =   300
      Width           =   1200
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   300
      Left            =   4200
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1185
   End
   Begin VB.CommandButton cmdSQL 
      Caption         =   "SQL &SELECT"
      Default         =   -1  'True
      Height          =   300
      Left            =   60
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2310
      Width           =   1230
   End
   Begin MSComctlLib.ImageList imlTreePics 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOSQLopzoeker.frx":0E54
            Key             =   "Tabel"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOSQLopzoeker.frx":0F66
            Key             =   "Opzoeking"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOSQLopzoeker.frx":1078
            Key             =   "Index"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOSQLopzoeker.frx":118A
            Key             =   "Eigenschap"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOSQLopzoeker.frx":129C
            Key             =   "Verbonden"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmADOSQLopzoeker.frx":13AE
            Key             =   "Kolom"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmbVersie 
      Caption         =   "&Versie"
      Height          =   300
      Left            =   6180
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1200
   End
   Begin MSComctlLib.ListView lvDatabase 
      Height          =   4395
      Left            =   7440
      TabIndex        =   7
      Top             =   0
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   7752
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "imlTreePics"
      SmallIcons      =   "imlTreePics"
      ColHdrIcons     =   "imlTreePics"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Index           =   1
      X1              =   60
      X2              =   9120
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   3
      Index           =   0
      X1              =   60
      X2              =   9120
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label lblRecordCount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6180
      TabIndex        =   10
      Top             =   0
      Width           =   1185
   End
End
Attribute VB_Name = "frmADOSQLOpzoeker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim grdColWidth(20) As Integer
Dim allesGesloten As Boolean

Dim datPrimaryRS As ADODB.Recordset
Dim querySQL As String

Private Sub cbOperatie_Click()

    queryPLUS
    queryChange

End Sub

Private Sub cbSQLBevel_Click()

    If cbSQLBevel.ListIndex <> 0 Then
        Msg = cbSQLBevel.text & " opdracht." & vbCr & vbCr & _
            "Het is ten zeerste aan te raden om dergelijke" & vbCr & _
            "opdrachten BINNENIN een TRANSACTIE uit te voeren" & vbCr & vbCr & vbCr & _
            "BEGIN WORK start een transactie" & vbCr & vbCr & _
            "ROLLBACK WORK annuleert alle wijziging na 'BEGIN WORK'" & vbCr & _
            "(m.a.w. zéér interessant om foutieve 'DELETE/UPDATES/INSERT'" & vbCr & _
            "opdrachten teniet te doen...)" & vbCr & vbCr & _
            "COMMIT WORK ten slotte laat alle 'DELETE/UPDATE/INSERT'" & vbCr & _
            "opdrachten doorgaan." & vbCr & vbCr & _
            "BEGIN WORK wordt hierna voorgesteld als instructie.  Druk Alt+E om te activeren"
            MsgBox Msg, vbExclamation
            txtSQL.text = "BEGIN WORK"
    Else
        queryPLUS
        queryChange
    End If
        
End Sub

Private Sub cbVelden_Click()

    queryPLUS
    queryChange
    
End Sub

Private Sub cmbSelect_Click()

bGet TABLE_VARIOUS, 1, "29" + cmbSelect.text
If Ktrl Then
Else
    RecordToVeld TABLE_VARIOUS
    If InStr(vBibTekst(TABLE_VARIOUS, "#v132 #"), "[Colwidth]") Then
        txtSQL.text = Left(vBibTekst(TABLE_VARIOUS, "#v132 #"), InStr(vBibTekst(TABLE_VARIOUS, "#v132 #"), "[Colwidth]") - 1)
    Else
        txtSQL.text = vBibTekst(TABLE_VARIOUS, "#v132 #")
    End If
    
    On Error Resume Next
    Msg = Mid(vBibTekst(TABLE_VARIOUS, "#v132 #"), InStr(vBibTekst(TABLE_VARIOUS, "#v132 #"), "[Colwidth]") + 10)
    If Msg = "" Then
        grdColWidth(0) = 0
    Else
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
    
End If

End Sub

Private Sub cmbSelect_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 46 Then
    bGet TABLE_VARIOUS, 1, "29" + cmbSelect.text
    If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "29" Then
    ElseIf MsgBox("Bestaande definitie '" + cmbSelect.text + "' verwijderen ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        RecordToVeld TABLE_VARIOUS
        bDelete TABLE_VARIOUS
        SelectComboVullen
    Else
        Exit Sub
    End If
ElseIf KeyCode = 13 Then
    cmdSelectWegschrijven_Click
End If

End Sub

Private Sub cmbVersie_Click()
Dim strVersionInfo As String

    On Error Resume Next
     strVersionInfo = "ADO Versie: " & adntDB.Version & vbCr & _
     "DBMS Naam: " & adntDB.Properties("DBMS Name") & vbCr & _
     "DBMS Versie: " & adntDB.Properties("DBMS Version") & vbCr & _
     "OLE DB Versie: " & adntDB.Properties("OLE DB Version") & vbCr & _
     "Provider Naam: " & adntDB.Properties("Provider Name") & vbCr & _
     "Provider Versie: " & adntDB.Properties("Provider Version") & vbCr
    MsgBox strVersionInfo
    
End Sub

Private Sub cmdBackup_Click()

    Dim TQString    As String
    Dim sTBLName    As String
    Dim SQLstring   As String
    
    Dim cnnSubHier As ADODB.Connection
    Dim rsSubHier As ADODB.Recordset
    Dim jetConnectSubHier As String
    Dim RecordTeller As Long
    Dim msgStringHier As String
    
    jetConnectSubHier = ADOJET_PROVIDER & _
            "Data Source=" & LOCATION_COMPANYDATA & _
            "\marnt.mdv;" & _
            "Persist Security Info=False"
    
    Set cnnSubHier = New ADODB.Connection
    cnnSubHier.Open jetConnectSubHier
    If Dir(LOCATION_COMPANYDATA & "\marnt.mdb") <> "" Then
        Kill LOCATION_COMPANYDATA & "\marnt.mdb"
    End If
    If Not CopyFile(PROGRAM_LOCATION, LOCATION_COMPANYDATA, "marnt.mdb") Then
        MsgBox "Stop" & vbCr & Error
    End If
    
    TQString = OpenSchemeAsString("TABLE")
    Do While TQString <> ""
        sTBLName = Mid(TQString, 1, InStr(TQString, vbCr) - 1)
        TQString = Mid(TQString, InStr(TQString, vbCr) + 1)
        SQLstring = "SELECT * INTO [" & LOCATION_COMPANYDATA & "marnt.mdb].[" & sTBLName & "] FROM " & sTBLName
        Err = 0
        On Error Resume Next
        SnelHelpPrint "Bezig aan tabel " & sTBLName, BL_LOGGING
        Screen.MousePointer = vbHourglass
        cnnSubHier.Execute SQLstring, RecordTeller
        If Err Then
            MsgBox Error
        Else
            msgStringHier = msgStringHier & RecordTeller & " records met succes overgedragen in tabel " & sTBLName & " / "
        End If
        Screen.MousePointer = vbNormal
    Loop
    MsgBox "Einde backup database" & vbCr & vbCr & msgStringHier, vbInformation
    MsgBox "Backup nog op veilige plaats bewaren: " & LOCATION_COMPANYDATA & "marnt.mdb", vbInformation
    MsgBox "Hierna worden de belangrijkste indexen nog aangemaakt in de marnt.mdb database"
    MsgBox "Stop"
    
    
End Sub

Private Sub cmdExecute_Click()

Dim recAantal As Long

On Error Resume Next
Screen.MousePointer = vbHourglass

Err = 0
adntDB.Execute txtSQL.text, recAantal
Screen.MousePointer = vbNormal
If Err Then
    MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
Else
    MsgBox txtSQL.text & vbCr & vbCr & "met succes uitgevoerd." & vbCr & vbCr & recAantal & " records met wijzigingen.", vbInformation
End If
'adoLoadDatabase

End Sub

Private Sub cmdKopij_Click()

Msg = "Kies 'Ja' voor kopij als XML bestand" & vbCr & _
    "Kies 'Nee' voor kopij naar het klassieke plakbord"

KtrlBox = MsgBox(Msg, vbQuestion + vbYesNoCancel + vbDefaultButton3)
If KtrlBox = vbCancel Then
    Exit Sub
ElseIf KtrlBox = vbNo Then
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText msfSQL.Clip
    Exit Sub
Else
    On Error GoTo CancelError
    Mim.Teken.filename = ""
    Mim.Teken.CancelError = True
    Mim.Teken.Filter = "Alle bestanden (*.xml)|*.xml"
    Mim.Teken.ShowSave
    Dim BestandHier As String
    BestandHier = Mim.Teken.filename
    If Not Dir(BestandHier) = "" Then
        Kill BestandHier
    End If
    datPrimaryRS.Save Mim.Teken.filename, adPersistXML
End If

CancelError:
Screen.MousePointer = vbNormal

End Sub

Private Sub cmdNet1_Click()

    Msg = "ALTER TABLE Journalen DROP COLUMN dece068"
    MsgBox "Alle rekening- en journaalvelden voor cijfers/bedragen dienen vanaf versie 6.5.301 bij voorkeur formaat DECIMAL te zijn (voorheen CURRENCY).  Wij zullen zolang als mogelijk een manuele hersamenstelling beschikbaar stellen d.m.v. de SQL instructie hierna EN VERVOLGENS opnieuw openen van het bedrijf.  Zeker tot aan versie 6.5.500 zal deze functie beschikbaar blijven." & vbCr & vbCr & "Aarzel NOOIT ons te contacteren voor bijkomende inlichtingen:" & vbCr & vbCr & Msg & vbCr & vbCr & "Uiteraard alléén indien U een verkeerde herrekening wenst te verbeteren..." & vbCr & vbCr & "NETWERKERS !  EERST AL UW ANDERE marINTEGRAAL GEBRUIKERSVERBINDINGEN STOPPEN !!!!", vbInformation

End Sub

Private Sub cmdOpenXML_Click()

    On Error GoTo CancelError
    Mim.Teken.filename = ""
    Mim.Teken.CancelError = True
    Mim.Teken.Filter = "Alle bestanden (*.xml)|*.xml"
    Mim.Teken.ShowOpen
    Dim BestandHier As String
    BestandHier = Mim.Teken.filename
    Err = 0
    On Error Resume Next
    datPrimaryRS.Close
    On Error Resume Next
    Err = 0
    datPrimaryRS.Open Mim.Teken.filename, , adOpenForwardOnly, adLockReadOnly, adCmdFile
    Set msfSQL.DataSource = datPrimaryRS
    If Err Then
        MsgBox "Dit is geen ADO-compatibel XML bestand.", vbInformation
    Else
        txtSQL = datPrimaryRS.Source
        lblRecordCount = datPrimaryRS.RecordCount
        msfSQL.Refresh
    End If
    
CancelError:
Screen.MousePointer = vbNormal

End Sub

Private Sub cmdSelectWegschrijven_Click()

bGet TABLE_VARIOUS, 1, "29" + cmbSelect.text
If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "29" Then
    'Als nieuw bijvoegen
    TLB_RECORD(TABLE_VARIOUS) = ""
    Msg = ""
    For COUNT_TO = 0 To msfSQL.Cols - 1
        Msg = Msg + Format(msfSQL.ColWidth(COUNT_TO)) + vbTab
    Next
    Msg = "[Colwidth]" + Msg
    Msg = txtSQL.text + Msg
    vBib TABLE_VARIOUS, Msg, "v132"
    vBib TABLE_VARIOUS, cmbSelect.text, "v250"
    vBib TABLE_VARIOUS, "29" + vBibTekst(TABLE_VARIOUS, "#v250 #"), "v005"
    bInsert TABLE_VARIOUS, 1
ElseIf MsgBox("Bestaande definitie '" + cmbSelect.text + "' overschrijven ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    RecordToVeld TABLE_VARIOUS
    Msg = ""
    For COUNT_TO = 0 To msfSQL.Cols - 1
        Msg = Msg + Format(msfSQL.ColWidth(COUNT_TO)) + vbTab
    Next
    Msg = "[Colwidth]" + Msg
    Msg = txtSQL.text + Msg
       
    vBib TABLE_VARIOUS, Msg, "v132"
    vBib TABLE_VARIOUS, cmbSelect.text, "v250"
    vBib TABLE_VARIOUS, "29" + vBibTekst(TABLE_VARIOUS, "#v250 #"), "v005"
    bUpdate TABLE_VARIOUS, 1
Else
    Exit Sub
End If
SelectComboVullen

End Sub

Private Sub cmdSluiten_Click()

Unload Me

End Sub

Private Sub cmdSQL_Click()

On Error Resume Next

lblRecordCount = ""
Me.Refresh
If adoRECORDset Then
    For COUNT_TO = 0 To msfSQL.Cols - 1
        If grdColWidth(COUNT_TO) = 0 Then
            Exit For
        Else
            msfSQL.ColWidth(COUNT_TO) = grdColWidth(COUNT_TO)
        End If
        msfSQL.ColAlignment(COUNT_TO) = flexAlignLeftTop
    Next
End If

End Sub


Private Sub Form_Load()

Set datPrimaryRS = New ADODB.Recordset
datPrimaryRS.CursorLocation = adUseClient

allesGesloten = False
Msg = "Moeten de tabellen gesloten worden (noodzakelijk om wijzigingen aan 'structuur' aan te brengen)"
If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
    MsgBox "De bestanden worden gesloten.  U kan eveneens wijzigingen aan de structuur van tabellin in de database aanbrengen.  Het is aanbevolen straks het bedrijf opnieuw te openen", vbInformation
    allesGesloten = True
    bClose 99
Else
    MsgBox "De bestanden worden niet gesloten.  U kan enkel de gegevens in de database bewerken.  U kan géén wijzigingen aan de structuur van tabellen in de database aanbrengen.", vbInformation
End If

Screen.MousePointer = vbHourglass

With msfSQL
    .Top = 5
    .Left = 0
    .Height = 2175
    .Width = 6135
End With
adoLoadDatabase

QueryNogEens:
bGetOrGreater TABLE_VARIOUS, 1, "29"
If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "29" Then
    'Steek er een voorbeeld in
          Msg = "SELECT TOP 6" + vbCrLf
    Msg = Msg + "    v019 AS RekNr," + vbCrLf
    Msg = Msg + "    v020 AS Omschrijving," + vbCrLf
    Msg = Msg + "    v022 AS [Boekjaar 0]," + vbCrLf
    Msg = Msg + "    v023 AS [Boekjaar -1]," + vbCrLf
    Msg = Msg + "    v024 AS [Boekjaar -2]," + vbCrLf
    Msg = Msg + "    v025 As [Boekjaar -3]" + vbCrLf
    Msg = Msg + "FROM" + vbCrLf
    Msg = Msg + "    Rekeningen" + vbCrLf
    Msg = Msg + "ORDER BY" + vbCrLf
    Msg = Msg + "    v020 DESC"
    TLB_RECORD(TABLE_VARIOUS) = ""
    vBib TABLE_VARIOUS, Msg, "v132"
    vBib TABLE_VARIOUS, "Query voorbeeld", "v250"
    vBib TABLE_VARIOUS, "29" + vBibTekst(TABLE_VARIOUS, "#v250 #"), "v005"
    bInsert TABLE_VARIOUS, 1
    If Ktrl Then
    Else
        GoTo QueryNogEens
    End If
Else
    SelectComboVullen
End If

Opzoeking:
cmdSQL_Click

On Error Resume Next
Screen.MousePointer = vbNormal

End Sub

Sub SelectComboVullen()

cmbSelect.Clear
bGetOrGreater TABLE_VARIOUS, 1, "29"
If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "29" Then
Else
    Do
        RecordToVeld TABLE_VARIOUS
        cmbSelect.AddItem vBibTekst(TABLE_VARIOUS, "#v250 #")
        bNext TABLE_VARIOUS
        If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "29" Then
            Exit Do
        End If
    Loop
End If
If cmbSelect.ListCount = 0 Then
Else
    cmbSelect.ListIndex = 0
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
datPrimaryRS.Close
Set datPrimaryRS = Nothing

'hier nu eerst terug rsjournaal maken !!!!!
If allesGesloten = True Then
    MsgBox "Bedrijfsdatabase wordt hierna automatisch afgesloten.", vbInformation
    AutoUnloadBedrijf
    Exit Sub
End If

End Sub

Private Sub lvDatabase_Click()

txtSQL = "SELECT * FROM " & lvDatabase.SelectedItem

End Sub

Private Sub lvDatabase_DblClick()

Dim COUNT_TO As Integer

txtSQL = "SELECT * FROM " & lvDatabase.SelectedItem
cmdSQL_Click
cbSQLBevel.Enabled = True
cbSQLBevel.ListIndex = 0

cbVelden.Enabled = True
cbVelden.Clear
For COUNT_TO = 0 To datPrimaryRS.fields.Count - 1
    cbVelden.AddItem datPrimaryRS.fields(COUNT_TO).Name
Next
If cbVelden.ListCount Then
    cbVelden.ListIndex = 0
    
End If
cbOperatie.Enabled = True
cbOperatie.ListIndex = 0
txtWaarde.Enabled = True
txtWaarde.text = "'%'"
txtWaarde.SetFocus

End Sub

Private Sub msfSQL_DblClick()

If msfSQL.Top = 5 Then
    'vergroten
    Me.WindowState = vbMaximized
    msfSQL.Top = 0
    msfSQL.Left = 0
    msfSQL.Height = Me.Height - 700
    msfSQL.Width = Me.Width - 100
Else
    'terug normaal
    Me.WindowState = vbNormal
    With msfSQL
    .Top = 5
    .Left = 0
    .Height = 2175
    .Width = 6135
    End With
End If

End Sub


Public Sub adoLoadDatabase()

    Dim itmX        As ListItem
    Dim TQString    As String
    Dim sTBLName    As String
    
    lvDatabase.ListItems.Clear
    TQString = OpenSchemeAsString("TABLE")

    'add the Tables
    Do While TQString <> ""
        sTBLName = Mid(TQString, 1, InStr(TQString, vbCr) - 1)
        TQString = Mid(TQString, InStr(TQString, vbCr) + 1)
        Set itmX = lvDatabase.ListItems.Add(, , sTBLName, "Tabel")
    Loop
    lvDatabase.View = lvwIcon
    
End Sub

Function adoRECORDset() As Boolean

    Screen.MousePointer = vbHourglass
    
    On Error Resume Next
    msfSQL.Recordset = Nothing
    
    datPrimaryRS.Close
    On Error Resume Next
    datPrimaryRS.Open txtSQL.text, adntDB
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
        msfSQL.Refresh
    Else
        Set msfSQL.DataSource = datPrimaryRS
    End If
    lblRecordCount = datPrimaryRS.RecordCount
    Screen.MousePointer = vbNormal

End Function

Private Sub txtPLUS_Change()
    
    queryChange
    
End Sub


Private Sub txtWaarde_Change()

    queryPLUS
    queryChange
    
End Sub

Function queryChange()

querySQL = cbSQLBevel.text & " " & txtPLUS.text & " " & cbVelden.text & " " & cbOperatie.text & " " & txtWaarde.text
txtSQL.text = querySQL

End Function

Function queryPLUS()

Select Case cbSQLBevel.ListIndex
    Case 0
        txtPLUS.text = " * FROM " & lvDatabase.SelectedItem & " WHERE "
        txtPLUS.Enabled = False
    
    Case 1
        txtPLUS.text = " FROM " & lvDatabase.SelectedItem & " WHERE "
        txtPLUS.Enabled = False
    
    Case 2
        txtPLUS.text = " " & lvDatabase.SelectedItem & " SET " & cbVelden.text & " = ??? WHERE "
        txtPLUS.Enabled = True
        
    Case Else
        MsgBox cbSQLBevel.text & " nog niet beschikbaar via snelinstructies", vbInformation
        txtSQL.text = ""
        Exit Function
End Select

End Function
