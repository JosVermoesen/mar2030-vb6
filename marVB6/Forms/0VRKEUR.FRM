VERSION 5.00
Begin VB.Form VoorkeurDefinities 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voorkeurinstellingen Editeerscherm (Fiches klanten, leveranciers, rekeningen)"
   ClientHeight    =   2430
   ClientLeft      =   690
   ClientTop       =   1185
   ClientWidth     =   8625
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
   ScaleHeight     =   2430
   ScaleWidth      =   8625
   Begin VB.ListBox VoorkeurVelden 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   4920
      TabIndex        =   2
      Top             =   810
      Width           =   3525
   End
   Begin VB.CommandButton DefEdit 
      Caption         =   "Toevoegen"
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
      Index           =   0
      Left            =   3630
      TabIndex        =   8
      Top             =   2010
      Width           =   1240
   End
   Begin VB.ListBox PredefVelden 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   30
      MultiSelect     =   2  'Extended
      TabIndex        =   6
      Top             =   810
      Width           =   3525
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
      Left            =   3630
      TabIndex        =   4
      Top             =   450
      Width           =   1240
   End
   Begin VB.CommandButton Definitie 
      Caption         =   "Definitie"
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
      Left            =   3630
      TabIndex        =   3
      Top             =   90
      Width           =   1240
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
      Left            =   870
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Gebruikers definitie"
      Height          =   195
      Index           =   7
      Left            =   4980
      TabIndex        =   5
      Top             =   600
      Width           =   1665
   End
   Begin VB.Label Label1 
      Caption         =   "Woordenboek definitie"
      Height          =   195
      Index           =   6
      Left            =   90
      TabIndex        =   7
      Top             =   570
      Width           =   1995
   End
   Begin VB.Label Label1 
      Caption         =   "Tabel"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   150
      Width           =   645
   End
End
Attribute VB_Name = "VoorkeurDefinities"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z


Private Sub DefEdit_Click(Index As Integer)

Select Case Index
    Case 0
        For i = 0 To PredefVelden.ListCount - 1
            If PredefVelden.Selected(i) Then
                VoorkeurVelden.AddItem PredefVelden.List(i)
            End If
        Next i
End Select

End Sub

Private Sub Definitie_Click()
Dim FlOutput As Integer

If VoorkeurVelden.ListCount = 0 Then
    Msg = "Gebruikerswoordenboek vernietigen." + vbCrLf + "Bent U zeker ?"
Else
    Msg = "Gebruikerswoordenboek bewaren." + vbCrLf + "Bent U zeker ?"
End If

    KtrlBox = MsgBox(Msg, 292)
    If KtrlBox = 6 Then
        If VoorkeurVelden.ListCount = 0 Then
            If Dir$(PROGRAM_LOCATION + "Def\" + Left(TABLEDEF_ONT(Val(Left(Tabel.Text, 1))), 3) + "U.Def") <> "" Then
                Kill PROGRAM_LOCATION + "Def\" + Left(TABLEDEF_ONT(Val(Left(Tabel.Text, 1))), 3) + "U.Def"
            End If
        Else
            FlOutput = FreeFile
            Open PROGRAM_LOCATION + "Def\" + Left(TABLEDEF_ONT(Val(Left(Tabel.Text, 1))), 3) + "U.Def" For Output As FlOutput
            For i = 0 To VoorkeurVelden.ListCount - 1
                Print #FlOutput, VoorkeurVelden.List(i); vbCrLf;
            Next i
            Close FlOutput
        End If
    Else
        Exit Sub
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
Tabel.ListIndex = 0

End Sub

Private Sub sluiten_Click()

Unload Me

End Sub

Private Sub Tabel_Click()
Dim T As Integer
Dim FlL As Integer

If TeleBibPagina(Val(Left(Tabel.Text, 1))) Then
Else
    MsgBox "Onjuiste situatie, kontakteer R&&Vsoft."
End If

End Sub

Private Function TeleBibPagina(Fl As Integer)
Dim FlInput As Integer
Dim FFDefinitie As String
Dim TBString As String
TeleBibPagina = False

On Local Error GoTo TeleBibError

'code= 1 - 1 : Poliskontrole 1 = ON
'      2 - 3 : Selektiekeuze waarde (ListIDX)
'      4 - 4 :
'      5 - 8 : TeleBib
'      9 - 9 : Volgnummer eventueel 1 tot 9
'      10    : * voor verplichte invulling
'Inlaadvolgorde 1 : 000     'hoofddefinitie (VSOFT)
'               2 : xxxG    'uitbreiding gebruiker
'               2 : xxxM    'uitbreiding makelaar

If Dir$(PROGRAM_LOCATION + "Def\" + Left(TABLEDEF_ONT(Fl), 3) + ".Def") = "" Then
    MsgBox "Geen TeleBib definitie " + Left(TABLEDEF_ONT(Fl), 3) + ".Def"
    Exit Function
End If

EerstEnVooral:
VoorkeurVelden.Clear
If Dir$(PROGRAM_LOCATION + "Def\" + Left(TABLEDEF_ONT(Fl), 3) + "U.Def") = "" Then
    GoTo GeenUserVoorkeur
Else
    FlInput = FreeFile
    Open PROGRAM_LOCATION + "Def\" + Left(TABLEDEF_ONT(Fl), 3) + "U.Def" For Input As FlInput
    Do While Not EOF(FlInput)
        Line Input #FlInput, TBString
        VoorkeurVelden.AddItem TBString
        T = T + 1
    Loop
    Close FlInput
End If

GeenUserVoorkeur:
PredefVelden.Clear
FlInput = FreeFile
Open PROGRAM_LOCATION + "Def\" + Left(TABLEDEF_ONT(Fl), 3) + ".Def" For Input As FlInput
T = 0
Do While Not EOF(FlInput)
    Line Input #FlInput, TBString
    PredefVelden.AddItem TBString
    T = T + 1
Loop
Close FlInput
TeleBibPagina = True

MakelaarIn:
If ProducentNummer = Space$(8) Then
ElseIf Dir$(PROGRAM_LOCATION + "Def\" + Left(TABLEDEF_ONT(Fl), 3) + "M.Def") = "" Then
Else
    FlInput = FreeFile
    Open PROGRAM_LOCATION + "Def\" + Left(TABLEDEF_ONT(Fl), 3) + "M.Def" For Input As FlInput
    Do While Not EOF(FlInput)
        Line Input #FlInput, TBString
        PredefVelden.AddItem TBString
    Loop
    Close FlInput
End If
Exit Function

TeleBibError:
MsgBox "Telebibinlaadfout" + Str$(T) + " error:" + Error$
Close FlInput
TeleBibPagina = False
Exit Function
Resume

End Function

Private Sub VoorkeurVelden_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 46
        Positie = VoorkeurVelden.ListIndex
        If Positie >= 0 Then
            VoorkeurVelden.RemoveItem Positie
        End If
End Select

End Sub

