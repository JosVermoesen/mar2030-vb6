VERSION 5.00
Begin VB.Form KwijtingEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gegevens Invoegen / Wijzigen"
   ClientHeight    =   3855
   ClientLeft      =   1815
   ClientTop       =   1605
   ClientWidth     =   7890
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
   LinkMode        =   1  'Source
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   7890
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox cbMail 
      Caption         =   "Voor Post && Mail"
      Height          =   375
      Left            =   1320
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3300
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CommandButton cbTB2 
      Caption         =   "Copy"
      Height          =   300
      Left            =   3660
      TabIndex        =   14
      Top             =   3360
      Width           =   1140
   End
   Begin VB.TextBox TekstInfo 
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
      Height          =   2085
      Index           =   5
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Top             =   1080
      Width           =   6420
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "Annuleren"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6600
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   420
      Width           =   1140
   End
   Begin VB.CommandButton Ok 
      Caption         =   "Ok"
      Height          =   300
      Left            =   6600
      TabIndex        =   4
      Top             =   60
      Width           =   1140
   End
   Begin VB.TextBox TekstInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   4410
      TabIndex        =   3
      Top             =   690
      Width           =   1560
   End
   Begin VB.TextBox TekstInfo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   1290
      TabIndex        =   2
      Top             =   720
      Width           =   1560
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1290
      TabIndex        =   5
      Top             =   30
      Width           =   4680
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   4410
      TabIndex        =   1
      Top             =   390
      Width           =   1560
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1290
      TabIndex        =   0
      Top             =   420
      Width           =   2280
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TB2"
      Height          =   240
      Index           =   5
      Left            =   540
      TabIndex        =   13
      Top             =   1080
      Width           =   570
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Loon"
      Height          =   240
      Index           =   4
      Left            =   3780
      TabIndex        =   11
      Top             =   720
      Width           =   570
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Premie"
      Height          =   255
      Index           =   3
      Left            =   510
      TabIndex        =   10
      Top             =   750
      Width           =   720
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datum"
      Height          =   285
      Index           =   2
      Left            =   3660
      TabIndex        =   9
      Top             =   390
      Width           =   690
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kliënt"
      Height          =   255
      Index           =   1
      Left            =   450
      TabIndex        =   8
      Top             =   60
      Width           =   795
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Polisnummer"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   7
      Top             =   450
      Width           =   1155
   End
End
Attribute VB_Name = "KwijtingEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Private Sub Annuleren_Click()

GridText = "ESC"
Unload Me

End Sub

Private Sub cbTB2_Click()

    MsgBox tb2Indent(TekstInfo(5).text)
    Clipboard.Clear
    Clipboard.SetText (TekstInfo(5).text)
    
    
End Sub

Private Sub Form_Load()
Dim astring As String

    KwijtingDrukken!PaperDetail.Col = 0
    TekstInfo(0).text = KwijtingDrukken!PaperDetail.text
    KwijtingDrukken!PaperDetail.Col = 1
    If KwijtingDrukken!PaperDetail.text = "" Then
        TekstInfo(2).text = KwijtingDrukken!TekstInfo(0).text
    Else
        TekstInfo(2).text = KwijtingDrukken!PaperDetail.text
    End If
    KwijtingDrukken!PaperDetail.Col = 2
    TekstInfo(3).text = KwijtingDrukken!PaperDetail.text
    KwijtingDrukken!PaperDetail.Col = 3
    TekstInfo(4).text = KwijtingDrukken!PaperDetail.text
    KwijtingDrukken!PaperDetail.Col = 4
    TekstInfo(1).text = KwijtingDrukken!PaperDetail.text
    KwijtingDrukken!PaperDetail.Col = 6
    TekstInfo(5).text = KwijtingDrukken!PaperDetail.text

End Sub

Private Sub Ok_Click()
Dim VeldPolis As String * 12

If Val(TekstInfo(3).text) + Val(TekstInfo(4).text) = 0 Then
    Beep
Else
    Dim astring As String
    
    KwijtingDrukken!PaperDetail.Col = 0
    KwijtingDrukken!PaperDetail.text = TekstInfo(0).text
    KwijtingDrukken!PaperDetail.Col = 1
    KwijtingDrukken!PaperDetail.text = TekstInfo(2).text
    KwijtingDrukken!PaperDetail.Col = 2
    KwijtingDrukken!PaperDetail.text = TekstInfo(3).text
    KwijtingDrukken!PaperDetail.Col = 3
    KwijtingDrukken!PaperDetail.text = TekstInfo(4).text
    KwijtingDrukken!PaperDetail.Col = 4
    KwijtingDrukken!PaperDetail.text = TekstInfo(1).text
        
    If cbMail.Value Then
    KwijtingDrukken!MailDetail.Col = 0
    KwijtingDrukken!MailDetail.text = TekstInfo(0).text
    KwijtingDrukken!MailDetail.Col = 1
    KwijtingDrukken!MailDetail.text = TekstInfo(2).text
    KwijtingDrukken!MailDetail.Col = 2
    KwijtingDrukken!MailDetail.text = TekstInfo(3).text
    KwijtingDrukken!MailDetail.Col = 3
    KwijtingDrukken!MailDetail.text = TekstInfo(4).text
    KwijtingDrukken!MailDetail.Col = 4
    KwijtingDrukken!MailDetail.text = TekstInfo(1).text
    KwijtingDrukken!cbMail.Enabled = True
    End If
    
    GridText = "OK"
    Unload Me
End If

End Sub


Private Sub TekstInfo_GotFocus(Index As Integer)

If Index = 1 Then SnelHelpPrint "Eén of meer tekens gevolgd door [Ctrl] om geïndexeerd te zoeken", BL_LOGGING
TekstInfo(Index).SelLength = Len(TekstInfo(Index).text)

End Sub

Private Sub TekstInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Dim msgTitel As String

If Index = 1 Then
    If KeyCode = 17 Then
        SharedFl = TABLE_CUSTOMERS
        aIndex = 1
        GridText = TekstInfo(1).text
        SqlSearch.Show 1
        If Ktrl Then
            TekstInfo(1).text = "-"
            KwijtingEdit!Ok.Enabled = False
            Exit Sub
        Else
            RecordToVeld TABLE_CUSTOMERS
            TekstInfo(1).text = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
            bGet TABLE_CONTRACTS, 1, vBibTekst(TABLE_CUSTOMERS, "#A110 #")
            If Ktrl Or vSet(KEY_BUF(TABLE_CONTRACTS), 12) <> vSet(vBibTekst(TABLE_CUSTOMERS, "#A110 #"), 12) Then
                MsgBox "Geen polissen voor deze klant te vinden !!"
                TekstInfo(1).text = "-"
                KwijtingEdit!Ok.Enabled = False
                Exit Sub
            Else
                Do
                    RecordToVeld TABLE_CONTRACTS
                    msgTitel = vBibTekst(TABLE_CONTRACTS, "#B010 #") & " " & vBibTekst(TABLE_CONTRACTS, "#e069 #")
                    
                    Msg = "Kwijting voor polisnummer : " + vBibTekst(TABLE_CONTRACTS, "#A000 #") + vbCrLf + vbCrLf
                    Msg = Msg + vBibTekst(TABLE_CONTRACTS, "#vs99 #") + vbCrLf
                    Msg = Msg + vBibTekst(TABLE_CONTRACTS, "#vs98 #") + vbCrLf + vbCrLf
                    Msg = Msg + "Bent U zeker ?"
                    KtrlBox = MsgBox(Msg, vbYesNo + vbDefaultButton2, msgTitel)
                    If KtrlBox = vbYes Then
                        TekstInfo(0).text = vBibTekst(TABLE_CONTRACTS, "#A000 #")
                        TekstInfo(0).SetFocus
                        TekstInfo(3).text = Val(vBibTekst(TABLE_CONTRACTS, "#B010 #"))
                        TekstInfo(4).text = Val(vBibTekst(TABLE_CONTRACTS, "#B014 #"))
                        Exit Do
                    End If
                    bNext TABLE_CONTRACTS
                    If Ktrl Or vSet(KEY_BUF(TABLE_CONTRACTS), 12) <> vSet(vBibTekst(TABLE_CUSTOMERS, "#A110 #"), 12) Then
                        MsgBox "Geen polissen meer voor deze klant te vinden !!"
                        TekstInfo(1).text = "-"
                        KwijtingEdit!Ok.Enabled = False
                        Exit Do
                    End If
                Loop
            End If
        End If
    End If
End If

End Sub



Private Sub TekstInfo_LostFocus(Index As Integer)

Select Case Index
    Case 0
        bGet TABLE_CONTRACTS, 0, vSet((TekstInfo(0).text), 12)
        If Ktrl Then
            KwijtingEdit!Ok.Enabled = False
            TekstInfo(1).SetFocus
            TekstInfo(0).text = ""
            Exit Sub
        Else
            RecordToVeld TABLE_CONTRACTS
            TekstInfo(0).text = vBibTekst(TABLE_CONTRACTS, "#A000 #")
            bGet TABLE_CUSTOMERS, 0, vBibTekst(TABLE_CONTRACTS, "#A110 #")
            If Ktrl Then
                If Left(KwijtingDrukken!KeuzeInfo(0).text, 4) <> vBibTekst(TABLE_CONTRACTS, "#A010 #") Then
                    TekstInfo(1).text = "KlantLink onmogelijk !!! Kontroleer !!!"
                    KwijtingEdit!Ok.Enabled = False
                End If
            Else
                RecordToVeld TABLE_CUSTOMERS
                TekstInfo(1).text = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
                If Left(KwijtingDrukken!KeuzeInfo(1).text, 1) = "1" Then
                    TekstInfo(2).text = Format(Val(vBibTekst(TABLE_CONTRACTS, "#v165 #")), "00") + "/" + Format(Val(vBibTekst(TABLE_CONTRACTS, "#v164 #")), "00") + "/" + Right(KwijtingDrukken!TekstInfo(0).text, 4)
                Else
                    TekstInfo(2).text = KwijtingDrukken!TekstInfo(0).text 'Format(Val(vBibTekst(TABLE_CONTRACTS, "#v165 #")), "00") + "/" + Format(Val(vBibTekst(TABLE_CONTRACTS, "#v164 #")), "00") + "/" + Right(KwijtingDrukken!TekstInfo(0).Text, 4)
                End If
                Ok.Enabled = True
            End If
        End If
        

    Case 2
        If DATE_INVALID((TekstInfo(2).text)) Then
            TekstInfo(2).text = KwijtingDrukken!TekstInfo(0).text
            Beep
            TekstInfo(2).SetFocus
        End If
    Case 3, 4
        TekstInfo(Index).text = Format(TekstInfo(Index).text, "#,##0.00")
End Select

End Sub

