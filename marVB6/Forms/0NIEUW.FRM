VERSION 5.00
Begin VB.Form NieuwBedrijf 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nieuw bedrijf"
   ClientHeight    =   1995
   ClientLeft      =   1845
   ClientTop       =   2415
   ClientWidth     =   6660
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
   HelpContextID   =   12000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1995
   ScaleWidth      =   6660
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtStartMaand 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1920
      TabIndex        =   17
      Text            =   "01/01"
      Top             =   1560
      Width           =   840
   End
   Begin VB.OptionButton TypeBoekjaar 
      Alignment       =   1  'Right Justify
      Caption         =   "Andere"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   840
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1560
      Width           =   915
   End
   Begin VB.TextBox Maanden 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   2820
      TabIndex        =   15
      Text            =   "12"
      Top             =   930
      Width           =   465
   End
   Begin VB.ComboBox CmbBedrijfsType 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      ItemData        =   "0NIEUW.frx":0000
      Left            =   3120
      List            =   "0NIEUW.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1560
      Width           =   3495
   End
   Begin VB.TextBox BedrijfsNaam 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   4995
   End
   Begin VB.CommandButton Installeren 
      Caption         =   "&Installeren"
      Height          =   360
      Left            =   5220
      TabIndex        =   3
      Top             =   30
      Width           =   1395
   End
   Begin VB.CommandButton Negeren 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   360
      Left            =   5220
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   450
      Width           =   1380
   End
   Begin VB.CheckBox Makelaar 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000016&
      Caption         =   "&Verzekeringsbemiddelaar"
      ForeColor       =   &H00000040&
      Height          =   240
      Left            =   4110
      TabIndex        =   4
      Top             =   960
      Width           =   2505
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2340
      TabIndex        =   5
      Top             =   0
      Width           =   2775
   End
   Begin VB.TextBox Boekjaar 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   930
      Width           =   840
   End
   Begin VB.OptionButton TypeBoekjaar 
      Alignment       =   1  'Right Justify
      Caption         =   "01/0&7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1020
      Width           =   795
   End
   Begin VB.OptionButton TypeBoekjaar 
      Alignment       =   1  'Right Justify
      Caption         =   "01/1&0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1320
      Width           =   795
   End
   Begin VB.OptionButton TypeBoekjaar 
      Alignment       =   1  'Right Justify
      Caption         =   "01/0&1 aanvangend"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Type"
      Height          =   285
      Index           =   1
      Left            =   3180
      TabIndex        =   14
      Top             =   1230
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Naam Bedrijf"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   9
      Top             =   30
      Width           =   585
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&DATA Lokatie"
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   12
      Top             =   30
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Boekjaar"
      Height          =   255
      Index           =   5
      Left            =   1920
      TabIndex        =   11
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aantal &maanden"
      Height          =   255
      Index           =   2
      Left            =   2820
      TabIndex        =   10
      Top             =   600
      Width           =   1485
   End
End
Attribute VB_Name = "NieuwBedrijf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim BedrijfsNummer As String
Dim Van As String

Dim FlN As Integer

Private Sub Bedrijfsnaam_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then SendKeys "{TAB}", False

End Sub

Private Function InstallTellers()

'Boekingen en algemene instellingen
SS99 "EUR", 296
SS99 Left(CmbBedrijfsType.text, 1), 20
SS99 "1", 299
SS99 "1", 306
SS99 "2", 21
SS99 "2", 291
SS99 "2", 290

SS99 "3", 183
SS99 "2", 200
SS99 "1", 201
SS99 "2", 290

'Aankoopverrichtingen
SS99 "00000", 1
SS99 "00000", 2
SS99 "00000", 3
SS99 "00000", 4
SS99 "00000", 15
SS99 "00000", 205

'Verkoopverrichtingen
SS99 "00000", 11
SS99 "00000", 12
SS99 "00000", 13
SS99 "00000", 14
SS99 "00000", 73
SS99 "00000", 59
SS99 "00000", 121
SS99 "2", 53
SS99 "0", 54
SS99 "1001", 181
SS99 "1", 182
SS99 "11", 185
SS99 "000", 186
SS99 "0", 187
SS99 "00000", 188
SS99 "2", 202
SS99 "1", 203
SS99 "2", 72
SS99 "1", 74
SS99 "1", 75
SS99 "1", 76
SS99 "0000.00", 300

'BTW Default Rekeningen
SS99 "498054", 16
SS99 "498055", 17
SS99 "498056", 18
SS99 "498057", 19
SS99 "498059", 22
SS99 "498063", 23
SS99 "498064", 24
SS99 "704000", 25
SS99 "604000", 77
SS99 "704100", 78
SS99 "340000", 79

'Default Collectieve Rekeningen
SS99 "400000", 9
SS99 "440000", 10
SS99 "4000", 297
SS99 "4400", 298
SS99 "756000", 27
SS99 "656000", 28
SS99 "489000", 145
SS99 "455", 146
SS99 "4899999", 147
SS99 "18", 148
SS99 "2999999", 149
SS99 "60", 150
SS99 "6089999", 151
SS99 "61", 152
SS99 "6559999", 153

'Financieel en Rekeningen
SS99 "571000", 41
SS99 " 0", 31
SS99 "561000", 42
SS99 " 0", 32
SS99 "551000", 43
SS99 " 0", 33
SS99 "552000", 44
SS99 " 0", 34
SS99 "553000", 45
SS99 " 0", 35
SS99 "582000", 39
SS99 " 0", 38
'SS99 "5540000", 211
'SS99 " 0", 215
'SS99 "5550000", 212
'SS99 " 0", 216
'SS99 "5560000", 213
'SS99 " 0", 217
'SS99 "5570000", 214
'SS99 " 0", 218
SS99 "551000", 101

'Bedrijfsinformatie
SS99 "MijnBedrijf bv", 46
SS99 "MijnStraatEnNummer", 47
SS99 "PCPC_MijnWoonplaats", 48
SS99 "Telefoon", 49
SS99 " ", 292
SS99 " ", 51
SS99 "IBAN", 293
SS99 "BIC", 294
SS99 "Mailadres Onderneming", 295
SS99 "Contactpersoon", 52
SS99 "Mailadres Contactpersoon", 50

'Status Boekjaar
SS99 "0", 62
SS99 "0", 63
SS99 "0", 64

'Kassaverkoop
SS99 "OFF", 130
SS99 "1", 131
SS99 "0", 132

End Function

Private Sub Boekjaar_LostFocus()

If Val(Boekjaar.text) < 2027 And Val(Boekjaar.text) > 2024 Then
    'ok
Else
    Boekjaar.text = Right(MIM_GLOBAL_DATE, 4)
    Boekjaar.SetFocus
    Beep
End If

End Sub


Private Sub Drive1_Change()
Dim A As String

Err = 0
On Local Error Resume Next
'ChDrive Left(Drive1.Drive, 1)
If Err Then MsgBox Error + vbCrLf + vbCrLf + "Kontroleer eerst a.u.b. !":  Exit Sub
Mid(LOCATION_, 1, 2) = Mid(Drive1.Drive, 1, 2)
A = Dir(LOCATION_ + "*.*")
If Err Then MsgBox Error + vbCrLf + vbCrLf + "Kontroleer eerst a.u.b. !":  Exit Sub
If A = "" Then MsgBox "Kontroleer eerst het opslagmedium a.u.b. via bedrijf openen !":  Exit Sub

BedrijfsNummer = VolgendBedrijf()
NieuwBedrijf.Caption = "Nieuw Bedrijf (" + BedrijfsNummer + ")"
'ChDrive Left(SYS_VAR(0), 2)
'ChDir SYS_VAR(0)
Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Top = 0
Left = 0

Unload BedrijfOpenen
Boekjaar.text = Right(MIM_GLOBAL_DATE, 4)
StatusPosition = 0

BedrijfsNummer = VolgendBedrijf()
NieuwBedrijf.Caption = "Nieuw Bedrijf (" + BedrijfsNummer + ")"

CmbBedrijfsType.AddItem "0: Standaard KMO"
CmbBedrijfsType.AddItem "1: marIntegraal NT Light"
CmbBedrijfsType.AddItem "2: Garage met Margeverkoop"
CmbBedrijfsType.AddItem "3: Auteursrechtadministratie"
CmbBedrijfsType.AddItem "4: Forfaitaire BTW"
CmbBedrijfsType.AddItem "5: Verzekeringsbemiddelaar"
CmbBedrijfsType.AddItem "6: Syndicus administratie"
CmbBedrijfsType.AddItem "7: BTW vrij(vb. verhuur onroerende goed, vzw)"
CmbBedrijfsType.AddItem "9: Gezinshuishouding"
CmbBedrijfsType.ListIndex = 0

TABLEDEF_ONT(TABLE_COUNTERS) = "00.ONT"                '09
bstNaam(TABLE_COUNTERS) = "jr" & NieuwBedrijf.Boekjaar.text

End Sub

Private Sub Installeren_Click()
Dim T As Integer
Dim X As Integer
Dim TotalNeeded As Long
Dim A As String

BAModus = 1
Van = Me.txtStartMaand + "/" + Boekjaar.text
A = "Installatie Bedrijf " + BedrijfsNummer + " voor :" + vbCrLf + vbCrLf + BedrijfsNaam.text + vbCrLf + vbCrLf + Format(Val(Maanden.text), "00") + " maanden vanaf : " + Van
If Makelaar.Value Then
    A = A + vbCrLf + vbCrLf + "Met programmakeuze voor makelaars !"
End If
Ktrl = MsgBox(A, 292)

If Not Ktrl = 6 Then
    Exit Sub
End If
    
TotalNeeded = 38024
destpath$ = LOCATION_ + BedrijfsNummer
If Not CreatePath(destpath$) Then
    MsgBox "foutmelding bij aanmaak " + destpath$
End If
    
LOCATION_COMPANYDATA = LOCATION_ + BedrijfsNummer + "\"

PeriodesMaken
LabelMaken
If Not CopyFile(App.path & "\Def", destpath$, "LICMARNT.###") Then GoTo ErrorSetup
If Not CopyFile(App.path & "\Def", destpath$, "DOK11.TXT") Then GoTo ErrorSetup
If Not CopyFile(App.path & "\Def", destpath$, "DOK12.TXT") Then GoTo ErrorSetup
If Not CopyFile(App.path & "\Def", destpath$, "DOK13.TXT") Then GoTo ErrorSetup
If Not CopyFile(App.path & "\Def", destpath$, "DOK14.TXT") Then GoTo ErrorSetup
If Not CopyFile(App.path & "\Def", destpath$, "DOK21.TXT") Then GoTo ErrorSetup
If Not CopyFile(App.path & "\Def", destpath$, "DOK22.TXT") Then GoTo ErrorSetup
If Not CopyFile(App.path & "\Def", destpath$, "DOK23.TXT") Then GoTo ErrorSetup
If Not CopyFile(App.path & "\Def", destpath$, "DOK24.TXT") Then GoTo ErrorSetup
If Not CopyFile(App.path, destpath$, "MARNT.MDV") Then GoTo ErrorSetup


jetConnect = ADOJET_PROVIDER & _
    "Data Source=" & LOCATION_COMPANYDATA & _
    "\marnt.mdv;" & _
    "Persist Security Info=False"

On Local Error Resume Next
bClose 99

Dim cat As New ADOX.Catalog
Dim tbl As New ADOX.Table
Dim Idx As New ADOX.Index

'Open the catalog.
cat.ActiveConnection = jetConnect
bstNaam(TABLE_COUNTERS) = "jr" & NieuwBedrijf.Boekjaar.text

tbl.Name = bstNaam(TABLE_COUNTERS)
    tbl.Columns.Append "v071", adVarWChar, 5
    tbl.Columns.Append "v217", adVarWChar, 30

' Define an index
With Idx
    .Name = "Setup Parameter"
    .Columns.Append "v071"
    .PrimaryKey = False
    .Unique = True
End With
    ' Append the index to the table
    tbl.Indexes.Append Idx
    
cat.Tables.Append tbl
cat.Tables.Refresh
Set cat = Nothing
Set tbl = Nothing
Set Idx = Nothing

On Local Error GoTo 0
Set adntDB = New ADODB.Connection

adntDB.Open jetConnect
Set ntDB = NTRuimte.OpenDatabase(LOCATION_COMPANYDATA + "marnt.MDV", False, False)
X = InitBestanden
X = InstallTellers
bClose 99
adntDB.Close
Set adntDB = Nothing
Negeren_Click

ExitSetup:
    Unload Me
    Exit Sub

ErrorSetup:
    MsgBox "Installatie nieuw bedrijf zonder succes.  Raadpleeg R&&VSOFT.", 48, dialogCaption$
    Unload Me
    Exit Sub

End Sub

Private Sub LabelMaken()

FlN = FreeFile
Open LOCATION_ + BedrijfsNummer + "/marnt.txt" For Output As FlN
    Print #FlN, BedrijfsNaam.text
Close FlN

End Sub

Private Sub Negeren_Click()

Unload Me

End Sub



Private Sub PeriodesMaken()
Dim b As String * 16
Dim A As String * 4
Dim T As Integer

aantalmaanden% = Val(Maanden.text)
VanafMaand% = Val(Mid(Van, 4, 2))
VanafJaar% = Val(Right(Van, 4))

Fl = FreeFile
Open PROGRAM_LOCATION + "9999.oct" For Random As Fl Len = 4
    A = "0"
        Put Fl, 1, A
    A = Right(Van, 4)
        Put Fl, 2, A
    A = "1"
        Put Fl, 3, A
Close Fl

Fl = FreeFile
Open LOCATION_ + BedrijfsNummer + "/def00.oct" For Random As Fl Len = 16
For T = 1 To aantalmaanden%
    b = Format(VanafJaar%, "0000") + Format(VanafMaand%, "00") + "01" + Format(VanafJaar%, "0000") + Format(VanafMaand%, "00") + Format(DAYS_IN_MONTH(VanafMaand%), "00")
    If VanafMaand% = 12 And T < aantalmaanden% Then
        VanafMaand% = 1
        VanafJaar% = VanafJaar% + 1
    Else
        VanafMaand% = VanafMaand% + 1
    End If
    Put Fl, T, b
Next T
b = ""

For t2% = T To 99
    Put Fl, t2%, b
Next
b = Format(VanafJaar%, "0000") + Format(VanafMaand% - 1, "00") + Format(DAYS_IN_MONTH(VanafMaand% - 1), "00") + Format(VanafJaar%, "0000") + Format(VanafMaand% - 1, "00") + Format(DAYS_IN_MONTH(VanafMaand% - 1), "00")
Put Fl, 99, b
Close Fl

End Sub

Private Function VolgendBedrijf$()
ReDim Aanwezig(99) As String
Dim T As Integer
Dim t2 As Integer
Dim EerstVolgendIs As String

On Local Error GoTo 0

For T = 1 To 999
    On Local Error Resume Next
    EerstVolgendIs = Dir$(LOCATION_ + Format(T, "000") + "\")
    If Err = 76 Or EerstVolgendIs = "" Then
        VolgendBedrijf = Format(T, "000")
        Exit Function
    End If
Next
MsgBox "Stop"
VolgendBedrijf = "000"

End Function

Private Sub txtStartMaand_LostFocus()

    If Len(txtStartMaand) <> 5 Then MsgBox "Enkel formaat dd/mm gebruiken a.u.b.": Me.txtStartMaand.text = "01/04": Me.txtStartMaand.SetFocus: Exit Sub
    If Mid(txtStartMaand, 3, 1) = "/" Then
    Else
        MsgBox "Enkel formaat dd/mm gebruiken a.u.b.": Me.txtStartMaand.text = "01/04": Me.txtStartMaand.SetFocus: Exit Sub
    End If
    
End Sub

Private Sub TypeBoekjaar_Click(Index As Integer)

    Select Case Index
            Case 0
                Me.txtStartMaand = "01/01"
                Me.txtStartMaand.Enabled = False
            Case 1
                Me.txtStartMaand = "01/07"
                Me.txtStartMaand.Enabled = False
            Case 2
                Me.txtStartMaand = "01/10"
                Me.txtStartMaand.Enabled = False
            Case 3
                Van = "01/04/" + Boekjaar.text
                Me.txtStartMaand = "01/04"
                Me.txtStartMaand.Enabled = True
                Me.txtStartMaand.SetFocus
    End Select
    
End Sub
