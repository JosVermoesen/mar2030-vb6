VERSION 5.00
Begin VB.Form frmVanGucht 
   Caption         =   "Volume- en maatberekening"
   ClientHeight    =   2550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2550
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   300
      Left            =   2430
      TabIndex        =   9
      Top             =   2160
      Width           =   1245
   End
   Begin VB.CommandButton cmdBijvoegen 
      Caption         =   "Lijn bijvoegen"
      Height          =   300
      Left            =   2400
      TabIndex        =   8
      Top             =   1290
      Width           =   1245
   End
   Begin VB.CommandButton cmdOvername 
      Caption         =   "&Overnemen"
      Enabled         =   0   'False
      Height          =   300
      Left            =   3330
      TabIndex        =   12
      Top             =   30
      Width           =   1245
   End
   Begin VB.TextBox txtHoogte 
      Alignment       =   1  'Right Justify
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
      Left            =   1380
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2160
      Width           =   975
   End
   Begin VB.TextBox txtBreedte 
      Alignment       =   1  'Right Justify
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
      Left            =   1380
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1860
      Width           =   975
   End
   Begin VB.TextBox txtLengte 
      Alignment       =   1  'Right Justify
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
      Left            =   1380
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtAantal 
      Alignment       =   1  'Right Justify
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
      Left            =   1380
      MaxLength       =   2
      TabIndex        =   1
      Top             =   1260
      Width           =   975
   End
   Begin VB.ListBox lstVolumeInfo 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label lblAantal 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Aantal"
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1260
      Width           =   1245
   End
   Begin VB.Label lblBreedte 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Breedte"
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1860
      Width           =   1245
   End
   Begin VB.Label lblLengte 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Lengte"
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   1245
   End
   Begin VB.Label lblHoogte 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Hoogte"
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   1245
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Volume Opnames"
      Height          =   255
      Left            =   210
      TabIndex        =   10
      Top             =   90
      Width           =   1635
   End
End
Attribute VB_Name = "frmVanGucht"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBijvoegen_Click()

Dim Bereken As Double

If Val(txtAantal) = 0 Then
    txtAantal.SetFocus
ElseIf Val(txtLengte) = 0 Then
    txtLengte.SetFocus
ElseIf Val(txtBreedte) = 0 Then
    txtBreedte.SetFocus
ElseIf Val(txtHoogte) = 0 Then
    txtHoogte.SetFocus
Else
    If Mid(WijzigenVerkoop.Keuze(0), 1, 1) = "5" Then
        Msg = "A: " + Dec(Val(txtAantal), "##") + " L: " + Dec(Val(txtLengte), "####") + " B: " + Dec(Val(txtBreedte), "####") + "        "
    Else
        Msg = "A: " + Dec(Val(txtAantal), "##") + " L: " + Dec(Val(txtLengte), "####") + " B: " + Dec(Val(txtBreedte), "####") + " H: " + Dec(Val(txtHoogte), "####")
    End If
    Bereken = txtLengte * txtBreedte * txtHoogte
    
    If Mid(WijzigenVerkoop.Keuze(0), 1, 1) = "5" Then
        If Bereken / 10000 < 0.1 Then
            Bereken = 0.1 * 10000
        Else
            Bereken = Bereken
        End If
    Else
        If Mid(WijzigenVerkoop.Keuze(0), 1, 1) = "3" Then
            If Bereken / 1000000 < 0.01 Then
                Bereken = 0.01 * 1000000
            Else
                Bereken = Bereken
            End If
        End If
    End If
    Bereken = Bereken * txtAantal
    If Mid(WijzigenVerkoop.Keuze(0), 1, 1) = "5" Then
        Bereken = CDbl(Format(Bereken / 10000, "#.###"))
    Else
        Bereken = CDbl(Format(Bereken / 1000000, "#.###"))
    End If
    Msg = Msg + " / " + Dec(Bereken, "#0.000")
    If Len(Msg) <> 38 Then
        MsgBox "Kontroleer de waarden a.u.b.  Max. grens werd overschreden"
        Exit Sub
    End If
    lstVolumeInfo.AddItem Msg
    txtAantal = ""
    txtLengte = ""
    txtBreedte = ""
    If Mid(WijzigenVerkoop.Keuze(0), 1, 1) = "5" Then
        txtHoogte = "0001"
    Else
        txtHoogte = ""
    End If
    cmdOvername.Enabled = True
End If

End Sub

Private Sub cmdOvername_Click()

Dim AlgTotaal As Double

XLogKey = ""
AlgTotaal = 0
For COUNT_TO = 0 To lstVolumeInfo.ListCount - 1
    If Len(lstVolumeInfo.List(COUNT_TO)) <> 38 Then
        MsgBox "Max. waarden overschreden, kontroleer eerst a.u.b."
        Exit Sub
    Else
        AlgTotaal = AlgTotaal + Val(Mid(lstVolumeInfo.List(COUNT_TO), 33))
        XLogKey = XLogKey + lstVolumeInfo.List(COUNT_TO) + vbCrLf
    End If
Next
WijzigenVerkoop.TekstInfo(6) = Dec(AlgTotaal, "#0.000")
Unload Me

End Sub

Private Sub cmdSluiten_Click()

GridText9 = ""
Unload Me

End Sub

Private Sub Form_Load()

If Mid(WijzigenVerkoop.Keuze(0), 1, 1) = "5" Then
    txtHoogte.Visible = False
    txtHoogte = Format(1, "0000")
    lblHoogte.Visible = False
End If

End Sub

Private Sub lstVolumeInfo_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode <> 46 Then Exit Sub

If lstVolumeInfo.ListCount <> 0 Then
    If lstVolumeInfo.ListIndex < 0 Then
    Else
        lstVolumeInfo.RemoveItem lstVolumeInfo.ListIndex
    End If
End If
If lstVolumeInfo.ListCount = 0 Then
    cmdOvername.Enabled = False
End If

End Sub

Private Sub txtAantal_LostFocus()

txtAantal = Format(txtAantal, "00")

End Sub

Private Sub txtBreedte_LostFocus()

txtBreedte = Format(txtBreedte, "0000")

End Sub

Private Sub txtHoogte_LostFocus()

txtHoogte = Format(txtHoogte, "0000")

End Sub

Private Sub txtLengte_LostFocus()

txtLengte = Format(txtLengte, "0000")

End Sub
