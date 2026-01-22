VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmADOXlog 
   AutoRedraw      =   -1  'True
   Caption         =   "Log"
   ClientHeight    =   2610
   ClientLeft      =   2880
   ClientTop       =   375
   ClientWidth     =   8445
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   21100
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2610
   ScaleWidth      =   8445
   Begin VB.CommandButton cmdDetail 
      Caption         =   "Selecteer Alles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3480
      TabIndex        =   6
      Top             =   2220
      Width           =   1275
   End
   Begin VB.CommandButton cmdOptie 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4800
      TabIndex        =   3
      Top             =   2220
      Visible         =   0   'False
      Width           =   1400
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      TabIndex        =   1
      Top             =   2220
      Width           =   555
   End
   Begin VB.CommandButton cmdSluiten 
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
      Height          =   330
      Left            =   2520
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2220
      Width           =   915
   End
   Begin MSComctlLib.ListView lvDetail 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648384
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
   Begin VB.Label lblAantal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   1
      Left            =   60
      TabIndex        =   5
      Top             =   2220
      Width           =   885
   End
   Begin VB.Label lblAantal 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Index           =   0
      Left            =   945
      TabIndex        =   4
      Top             =   2220
      Width           =   885
   End
End
Attribute VB_Name = "frmADOXlog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDetail_Click()

    Dim T As Integer
    
        GridText = ""
        For T = 1 To lvDetail.ListItems.Count
            lvDetail.ListItems(T).Checked = True
            GridText = GridText + lvDetail.ListItems(T).SubItems(1) + vbCr
        Next
        
End Sub

Private Sub cmdOK_Click()

    XLogKey = lvDetail.SelectedItem + vbCrLf
        XLogKey = XLogKey + lvDetail.SelectedItem.SubItems(1)
    Me.Hide

End Sub

Private Sub cmdSluiten_Click()

    XLogKey = ""
    GridText = ""
    Me.WindowState = 0
    Me.Hide
    
End Sub

Private Sub cmdOptie_Click()
Dim T As Integer

    Msg = lblAantal(1) & " van " & lblAantal(0) & " lijnen geselecteerd voor " & cmdOptie.Caption & vbCr & vbCr
    Msg = Msg + "Bent U zeker"
    KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
    If KtrlBox = vbYes Then
        GridText = ""
        For T = 1 To lvDetail.ListItems.Count
            If cmdOptie.Tag = "Kassa" Then
                If lvDetail.ListItems(T).Checked Then
                    GridText = GridText & lvDetail.ListItems(T) & vbCr
                End If
            Else
                If lvDetail.ListItems(T).Checked Then
                    GridText = GridText + lvDetail.ListItems(T).SubItems(1) + vbCr
                End If
            End If
        Next
        Me.Hide
    End If

End Sub

Private Sub Form_Load()

    LaadInstellingen Me
    Me.Width = 9525
    Me.Height = 6525

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    lvDetail.Top = 0
    lvDetail.Left = 0
    lvDetail.Width = Me.Width - 100
    lvDetail.Height = Me.Height - 1200

    cmdSluiten.Top = Me.Height - 1100
    cmdOK.Top = Me.Height - 1100
    cmdOptie.Top = Me.Height - 1100
    cmdDetail.Top = Me.Height - 1100
    lblAantal(0).Top = Me.Height - 1100
    lblAantal(1).Top = Me.Height - 1100
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Boolean

X = InstellingenBewaard(Me)

End Sub

Private Sub lvDetail_ItemCheck(ByVal Item As MSComctlLib.ListItem)

If Item.Checked Then
    lblAantal(1) = Val(lblAantal(1)) + 1
Else
    lblAantal(1) = Val(lblAantal(1)) - 1
End If
If lblAantal(1) = 0 Then
    cmdOptie.Visible = False
    cmdDetail.Visible = False
Else
    cmdOptie.Visible = True
    If cmdDetail.Tag = "DetailLijnen" Then
        cmdDetail.Visible = True
    End If
End If


End Sub

Private Sub lvDetail_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 46
        Msg = "dokument [" + Trim$(lvDetail.SelectedItem.SubItems(3)) + "] verwijderen."
        KtrlBox = MsgBox(Msg, 292)
        If KtrlBox = vbYes Then
            MsgBox "stop"
        End If
End Select
lvDetail.SetFocus

End Sub

