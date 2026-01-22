VERSION 5.00
Begin VB.Form ProgrammaStruktuur 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Systeem-Definitie raadplegen/wijzigen"
   ClientHeight    =   1755
   ClientLeft      =   1530
   ClientTop       =   1965
   ClientWidth     =   6480
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
   ScaleHeight     =   1755
   ScaleWidth      =   6480
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Update2 
      Caption         =   "Wijzigen via Kladblok"
      Enabled         =   0   'False
      Height          =   345
      Left            =   4080
      TabIndex        =   2
      Top             =   0
      Width           =   2325
   End
   Begin VB.ComboBox DefinitieBestand 
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   3045
   End
   Begin VB.ListBox DefinitieTekst 
      Columns         =   4
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   1080
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   420
      Width           =   5355
   End
   Begin VB.Label Label1 
      Caption         =   "Lengte"
      Height          =   225
      Index           =   3
      Left            =   150
      TabIndex        =   6
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "VarType"
      Height          =   225
      Index           =   2
      Left            =   150
      TabIndex        =   5
      Top             =   870
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "InfoTekst"
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   4
      Top             =   660
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "EdibiCode"
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   450
      Width           =   1005
   End
End
Attribute VB_Name = "ProgrammaStruktuur"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub DefinitieBestand_Click()
Dim FlTemp As Integer
Dim TekstString As String

On Local Error GoTo EndInlaad

DefinitieTekst.Clear

FlTemp = FreeFile
Open PROGRAM_LOCATION + "Def\" + DefinitieBestand.Text For Input As FlTemp
While Not EOF(FlTemp)
    Input #FlTemp, TELEBIB_CODE(0), TELEBIB_TEXT(0), TELEBIB_TYPE(0), TELEBIB_LENGTH(0)
    DefinitieTekst.AddItem TELEBIB_CODE(0)
    DefinitieTekst.AddItem TELEBIB_TEXT(0)
    DefinitieTekst.AddItem TELEBIB_TYPE(0)
    DefinitieTekst.AddItem Format(TELEBIB_LENGTH(0), "000")
Wend
Close FlTemp
If DefinitieTekst.ListCount Then DefinitieTekst.ListIndex = 0
Exit Sub

EndInlaad:
Close FlTemp
Exit Sub

End Sub

Private Sub DefinitieBestand_GotFocus()

Update2.Enabled = True
Update2.Default = True

End Sub

Private Sub DefinitieBestand_LostFocus()

'Update2.Enabled = False
'Update2.Default = False

End Sub

Private Sub Form_Load()
Dim Bestand As String

Bestand = Dir$(PROGRAM_LOCATION + "Def\" + "*.Def")
Do While Bestand <> ""
    DefinitieBestand.AddItem Bestand
    Bestand = Dir$
Loop
If DefinitieBestand.ListCount Then DefinitieBestand.ListIndex = 0

End Sub

Private Sub Update2_Click()
Dim X As Integer

'On Error Resume Next
'Shell "notepad.exe " + HetBESTAND, vbNormalFocus
X = Shell("notepad.exe " + PROGRAM_LOCATION + "Def\" + DefinitieBestand.Text, 1)

End Sub

