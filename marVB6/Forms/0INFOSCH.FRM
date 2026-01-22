VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form InfoScherm 
   Caption         =   "InfoScherm"
   ClientHeight    =   2625
   ClientLeft      =   840
   ClientTop       =   4095
   ClientWidth     =   8700
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2625
   ScaleWidth      =   8700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdSluiten 
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
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton Speciaal 
      Caption         =   "Speciaal"
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
      Left            =   1200
      TabIndex        =   0
      Top             =   2220
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid X 
      Height          =   2055
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6585
      _ExtentX        =   11615
      _ExtentY        =   3625
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   -2147483624
      SelectionMode   =   1
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
End
Attribute VB_Name = "InfoScherm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z




Private Sub cmdSluiten_Click()

Unload Me

End Sub

Private Sub Form_Load()

If VSF_PRO Then
    Speciaal.Visible = True
End If

End Sub

Private Sub Form_Resize()

On Local Error Resume Next

    X.Top = 0
    X.Left = 0
    X.Width = InfoScherm.Width - 100
    X.Height = InfoScherm.Height - 1000
    Speciaal.Top = InfoScherm.Height - 870
    cmdSluiten.Top = InfoScherm.Height - 870

End Sub

Private Sub Speciaal_Click()

Msg = "Alle journaallijnen van dit dokument vernietigen en journaal opnieuw opbouwen met de overblijvende lijnen.  Bent U zeker ?"
Ktrl = MsgBox(Msg, 292)
If Ktrl = 6 Then
    Screen.MousePointer = vbHourglass
    Do
        bGet TABLE_JOURNAL, 1, Right(InfoScherm.Caption, 8)
        If Ktrl Then
            Exit Do
        Else
            bDelete TABLE_JOURNAL
            If Ktrl Then
                MsgBox "Stop in verwijderen journaallijn"
            End If
        End If
    Loop

    X.Col = 7
    For COUNT_TO = 1 To X.Rows - 1
        X.Row = COUNT_TO
        If X.Text = "" Then Exit For
        TLB_RECORD(TABLE_JOURNAL) = X.Text
        bInsert TABLE_JOURNAL, 0
        If Ktrl Then
            MsgBox "Stop in wegboeken journaal !"
        End If
    Next
End If
Screen.MousePointer = vbNormal

End Sub

Private Sub X_DblClick()

Dim MijnInput As Variant
Dim TempoX As Integer
Dim TempoY As Integer
Dim vsbCodeHier As String

If Speciaal.Visible = True Then
    If X.Col = 7 Then
        MsgBox "Stop"
        
    Else
        MijnInput = InputBox("", , X.Text)
        If Len(MijnInput) <> 0 Then
            X.Text = MijnInput
            TempoX = X.Col
            TempoY = X.Row
            X.Col = 7
            TLB_RECORD(TABLE_JOURNAL) = X.Text
            
            X.Col = TempoX
            X.Row = 0
            vsbCodeHier = Mid(X.Text, InStr(X.Text, "#") + 1)
            vBib TABLE_JOURNAL, (MijnInput), vsbCodeHier
                                    
            X.Row = TempoY
            X.Col = 7
            X.Text = TLB_RECORD(TABLE_JOURNAL)
            X.Col = TempoX
        End If
    End If
End If

End Sub

Private Sub X_KeyDown(KeyCode As Integer, Shift As Integer)

On Local Error Resume Next

If KeyCode = 46 And Speciaal.Visible = True Then
    X.RemoveItem X.Row
End If

End Sub

Private Sub X_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
    Unload Me
End If

End Sub

