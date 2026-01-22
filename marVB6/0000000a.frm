VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form AlfaSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Geïndexeerd Zoeken"
   ClientHeight    =   3375
   ClientLeft      =   375
   ClientTop       =   1995
   ClientWidth     =   8760
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
   ScaleHeight     =   3375
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Zoeken 
      Caption         =   "Zoeken"
      Default         =   -1  'True
      Height          =   360
      Left            =   7440
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton ExitToets 
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
      Height          =   360
      Left            =   7440
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.ComboBox Sortering 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   480
      Width           =   3720
   End
   Begin MSFlexGridLib.MSFlexGrid Lijst 
      Height          =   2415
      Left            =   30
      TabIndex        =   6
      Top             =   870
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   4260
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   16777152
      FocusRect       =   2
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
      Caption         =   "Te &Zoeken"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Sortering op"
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "AlfaSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim IndexNR As Integer


Private Sub ExitToets_Click()

Ktrl = 99
Unload Me

End Sub




Private Sub Form_Load()
Dim T As Integer
Dim SorteringTel As Integer

Mim!SnelHelpTijd.Enabled = False

Lijst.Cols = ntDB.TableDefs(BstNaam(SharedFl)).Indexes.Count
Lijst.Rows = 2
Lijst.Row = 0

SorteringTel = 0
Caption = Caption + " " + BstNaam(SharedFl)
For T = 0 To ntDB.TableDefs(BstNaam(SharedFl)).Indexes.Count - 1
    With Lijst
        .Col = T
        .Text = FLIndexCaption(SharedFl, T)
        If T = 1 Then
            .ColWidth(T) = 3945
        Else
            .ColWidth(T) = 1650
        End If
        .ColAlignment(T) = flexAlignLeftTop
    End With
    Sortering.AddItem ntDB.TableDefs(BstNaam(SharedFl)).Indexes(T).Name
    'Debug.Print NTDB.TableDefs(BstNaam(SharedFl)).Indexes(T).Name, FLIndexCaption(SharedFl, T)
    If ntDB.TableDefs(BstNaam(SharedFl)).Indexes(T).Name = FLIndexCaption(SharedFl, aIndex) Then
        SorteringTel = T
        IndexNR = T
    End If
Next
Sortering.ListIndex = SorteringTel
If InStr(GridText, "@Beperk@") Then
    Text1.Text = Left$(GridText, 2)
    Zoeken_Click
ElseIf GridText <> "" Then
    Text1.Text = GridText
    Zoeken_Click
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Mim!SnelHelpTijd.Enabled = True

End Sub

Private Sub Lijst_Click()

On Local Error Resume Next
Dim DezePipo As String
Dim SnelString As String

DezePipo = Lijst.TextMatrix(Lijst.Row, 0)
SnelString = ""
bGet SharedFl, 0, DezePipo
If Ktrl Then
    SnelHelpPrint "~"
    Exit Sub
Else
    RecordToVeld SharedFl
End If
On Local Error Resume Next
Err = 0
For TelTot = 0 To 30
    SnelString = SnelString + Trim(vBibTekst(SharedFl, "#" + vBC(SharedFl, TelTot) + " #")) + " "
    If Err Then Exit For
Next
SnelHelpPrint "~" + Trim(SnelString)
Text1.Text = Lijst.TextMatrix(Lijst.Row, IndexNR)

End Sub

Private Sub Lijst_DblClick()

Zoeken_Click

End Sub



Private Sub Lijst_RowColChange()

If Lijst.Rows <> 2 Then
    Lijst_Click
End If

End Sub

Private Sub SorTering_Click()
Dim T As Integer

IndexNR = -1
For T = 0 To Sortering.ListCount - 1
    If Sortering.Text = FLIndexCaption(SharedFl, T) Then
        IndexNR = T
    End If
Next
If IndexNR < 0 Then
    MsgBox "Indexen zijn ondertussen vernieuwd ?" + vbCrLf + vbCrLf + "Om van de nieuwe indexen gebruik te kunnen maken dient U het bedrijf te heropenen !"
End If

End Sub






Private Sub Text1_Change()

If RTrim(Text1.Text) = "" Then
    Zoeken.Caption = "Zoeken"
ElseIf Text1.Text = Lijst.Text Then
    Zoeken.Caption = "Ok"
Else
    Zoeken.Caption = "Zoeken"
End If

End Sub

Private Sub Text1_GotFocus()

Text1.SelLength = Len(Text1.Text)
SnelHelpPrint "Minstens 1 teken om te zoeken (+ ENTER).  Maak daarna een keuze uit de lijst via de pijltjestoets OMHOOG of OMLAAG"

End Sub


Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 38 To 40
        Lijst.SetFocus
        Lijst.Col = IndexNR
        Lijst.Row = 1

End Select

End Sub

Private Sub Zoeken_Click()
Dim KeuzeIndex  As Integer
Dim Tel As Integer

Dim T As Integer
Dim a As String
Dim aa As String
Dim Sleuteltje As String

On Local Error Resume Next

If Zoeken.Caption = "Ok" Then
    Lijst.Col = 0
    bGet SharedFl, 0, Lijst.Text
    If Ktrl Then
        MsgBox "..."
    Else
        RecordToVeld SharedFl
    End If
    Unload Me
    Exit Sub
End If

If Len(Text1.Text) < 1 Then
    If AlfaSearch.Visible Then
        MsgBox "Minstens één teken als zoek-criteria a.u.b. !"
        Text1.SetFocus
        Exit Sub
    End If
End If

Lijst.Rows = 2
Lijst.Row = 1
For Tel = 0 To Sortering.ListCount
    Lijst.Col = Tel
    Lijst.Text = ""
Next
Lijst.Refresh
KeuzeIndex = IndexNR
Sleuteltje = UCase$(Text1.Text)
bGetOrGreater SharedFl, KeuzeIndex, (Text1.Text)
If Ktrl Or UCase(Left(KeyBuf(SharedFl), Len(Sleuteltje))) > UCase(Sleuteltje) Then
Else
    Screen.MousePointer = 11
    RecordToVeld SharedFl
    Text1.Text = vBibTekst(SharedFl, "#" + Left(FlIndexIs(SharedFl, KeuzeIndex), 4) + " #")
    aa = ""
    For Tel = 0 To FlAantalIndexen(SharedFl)
        aa = aa + RTrim(vBibTekst(SharedFl, "#" + Left(FlIndexIs(SharedFl, Tel), 4) + " #")) + vbTab
    Next
    Lijst.AddItem aa, Lijst.Rows - 1
    T = 1
    Do While Ktrl = 0
        If Asc(Sleuteltje) = 130 Then Exit Do
        bNext SharedFl
        If Ktrl Or UCase(Left(KeyBuf(SharedFl), Len(Sleuteltje))) > UCase(Sleuteltje) Then
            Exit Do
        Else
            RecordToVeld SharedFl
            aa = ""
            For Tel = 0 To FlAantalIndexen(SharedFl)
                aa = aa + RTrim(vBibTekst(SharedFl, "#" + Left(FlIndexIs(SharedFl, Tel), 4) + " #")) + vbTab
            Next
            Lijst.AddItem aa, Lijst.Rows - 1
            SnelHelpPrint RTrim$(vBibTekst(SharedFl, "#" + FlIndexIs(SharedFl, 0) + "#")) + " " + RTrim$(vBibTekst(SharedFl, "#" + FlIndexIs(SharedFl, 1) + "#"))
        End If
    Loop
    Lijst.Row = 1
    Zoeken.Caption = "Ok"
    Screen.MousePointer = 0
    Lijst.SetFocus
    Lijst.Col = IndexNR
    Lijst.Row = 1
End If

If aa = "" Then
    Text1.Text = ""
    Text1.SetFocus
    Exit Sub
Else
    On Local Error Resume Next
    Lijst.Row = 1
    Lijst.SetFocus
End If

End Sub

