VERSION 5.00
Begin VB.Form MijStat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Toepassingen Maatschappijen (AS/1 verzoeken)"
   ClientHeight    =   3645
   ClientLeft      =   1920
   ClientTop       =   1275
   ClientWidth     =   6090
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
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3645
   ScaleWidth      =   6090
   Begin VB.ListBox MijDetail 
      BackColor       =   &H00C0C0C0&
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
      Height          =   2595
      Left            =   90
      TabIndex        =   2
      Top             =   420
      Width           =   5925
   End
   Begin VB.ComboBox KeuzeInfo 
      BackColor       =   &H00C0C0C0&
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
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   30
      Width           =   5895
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   390
      Left            =   4440
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3150
      Width           =   1560
   End
End
Attribute VB_Name = "MijStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim AS1Mij(90) As String

Private Sub Annuleren_Click()

Unload Me

End Sub

Private Sub Form_Load()

InitASBestanden

End Sub


Private Sub InitASBestanden()
Dim T As String
Dim FunktieT As String
Dim TT As Integer
Dim FlTemp As Integer

On Local Error GoTo InitHandler

FlTemp = FreeFile
Open PROGRAM_LOCATION + "Def\" + "as1oct.ini" For Input As FlTemp
Line Input #FlTemp, T
Line Input #FlTemp, T
For TT = 0 To Val(T)
    Line Input #FlTemp, Dummy$
Next

Line Input #FlTemp, T
Line Input #FlTemp, T
For TT = 0 To Val(T)
    Line Input #FlTemp, Dummy$
Next
Line Input #FlTemp, T
Line Input #FlTemp, T
For TT = 0 To Val(T)
    Line Input #FlTemp, FunktieT
    If FunktieT = "" Then
    Else
        KeuzeInfo(0).AddItem FunktieT
    End If
Next
KeuzeInfo(0).ListIndex = 0
Close 1
Exit Sub

InitHandler:
MsgBox "ASOCT.INI beschadigd of niet aanwezig.  Installer het programma opnieuw !"
End

End Sub


Private Sub KeuzeInfo_Click(Index As Integer)

If Index <> 0 Then Exit Sub
LogKontrole

End Sub


Private Sub LogKontrole()

MijDetail.Clear

bGetOrGreater TABLE_VARIOUS, 1, "23CO"
If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 4) <> "23CO" Then
    MsgBox "Toepassingen maatschappijen eerst inladen a.u.b."
    Exit Sub
End If

Do
    RecordToVeld TABLE_VARIOUS
    If Mid(vBibTekst(TABLE_VARIOUS, "#v220 #"), Val(Left(KeuzeInfo(0).Text, 3)), 1) = "1" Then
        bGet TABLE_SUPPLIERS, 0, Mid(FVT(TABLE_VARIOUS, 0), 2)
        If Ktrl Then
            MijDetail.AddItem Mid(FVT(TABLE_VARIOUS, 0), 4, 4) + ": " + RTrim$(vBibTekst(TABLE_VARIOUS, "#A102 #")) + ", (geen agentschap)"
        Else
            RecordToVeld TABLE_SUPPLIERS
            MijDetail.AddItem Mid(FVT(TABLE_VARIOUS, 0), 4, 4) + ": " + RTrim$(vBibTekst(TABLE_VARIOUS, "#A102 #")) + ", agentnr. " + RTrim$(vBibTekst(TABLE_SUPPLIERS, "#A100 #")) + " : " + vBibTekst(TABLE_SUPPLIERS, "#A400 #")
        End If
    End If
    bNext TABLE_VARIOUS
    If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 4) <> "23CO" Then
        Exit Do
    End If
Loop

End Sub


