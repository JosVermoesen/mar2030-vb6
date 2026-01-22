VERSION 5.00
Begin VB.Form BoekKwijtingEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gegevens Invoegen / Wijzigen"
   ClientHeight    =   2070
   ClientLeft      =   1815
   ClientTop       =   1605
   ClientWidth     =   5340
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
   ScaleHeight     =   2070
   ScaleWidth      =   5340
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   1560
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "Negeren"
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
      Left            =   3990
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   750
      Width           =   1240
   End
   Begin VB.CommandButton Ok 
      Caption         =   "Ok"
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
      Left            =   3990
      TabIndex        =   5
      Top             =   390
      Width           =   1240
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   1290
      TabIndex        =   3
      Top             =   1350
      Width           =   1560
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   1290
      TabIndex        =   2
      Top             =   1020
      Width           =   1560
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1290
      TabIndex        =   6
      Top             =   30
      Width           =   3960
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1290
      TabIndex        =   1
      Top             =   690
      Width           =   1560
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1290
      TabIndex        =   0
      Top             =   360
      Width           =   2280
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Beheerskost"
      Height          =   240
      Index           =   5
      Left            =   60
      TabIndex        =   13
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Loon"
      Height          =   240
      Index           =   4
      Left            =   660
      TabIndex        =   12
      Top             =   1380
      Width           =   570
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Premie"
      Height          =   255
      Index           =   3
      Left            =   510
      TabIndex        =   11
      Top             =   1050
      Width           =   720
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datum"
      Height          =   285
      Index           =   2
      Left            =   540
      TabIndex        =   10
      Top             =   690
      Width           =   690
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kliënt"
      Height          =   255
      Index           =   1
      Left            =   450
      TabIndex        =   9
      Top             =   60
      Width           =   795
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Polisnummer"
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   390
      Width           =   1155
   End
End
Attribute VB_Name = "BoekKwijtingEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Private Sub Annuleren_Click()

GridText = "ESC"
Unload Me

End Sub

Private Sub Form_Load()
Dim astring As String

    KwijtingBoeken!PolisDetail.Col = 0
    TekstInfo(0).Text = KwijtingBoeken!PolisDetail.Text
    KwijtingBoeken!PolisDetail.Col = 1
    If KwijtingBoeken!PolisDetail.Text = "" Then
        TekstInfo(2).Text = KwijtingBoeken!TekstInfo(0).Text
    Else
        TekstInfo(2).Text = KwijtingBoeken!PolisDetail.Text
    End If
    KwijtingBoeken!PolisDetail.Col = 2
    TekstInfo(3).Text = KwijtingBoeken!PolisDetail.Text
    KwijtingBoeken!PolisDetail.Col = 3
    TekstInfo(4).Text = KwijtingBoeken!PolisDetail.Text
    KwijtingBoeken!PolisDetail.Col = 4
    TekstInfo(1).Text = KwijtingBoeken!PolisDetail.Text
    KwijtingBoeken!PolisDetail.Col = 6
    TekstInfo(5).Text = KwijtingBoeken!PolisDetail.Text

End Sub

Private Sub Ok_Click()
Dim VeldPolis As String * 12

If Val(TekstInfo(3).Text) + Val(TekstInfo(4).Text) = 0 Then
    Beep
Else
    Dim astring As String
    
    KwijtingBoeken!PolisDetail.Col = 0
    KwijtingBoeken!PolisDetail.Text = TekstInfo(0).Text
    KwijtingBoeken!PolisDetail.Col = 1
    KwijtingBoeken!PolisDetail.Text = TekstInfo(2).Text
    KwijtingBoeken!PolisDetail.Col = 2
    KwijtingBoeken!PolisDetail.Text = TekstInfo(3).Text
    KwijtingBoeken!PolisDetail.Col = 3
    KwijtingBoeken!PolisDetail.Text = TekstInfo(4).Text
    KwijtingBoeken!PolisDetail.Col = 4
    KwijtingBoeken!PolisDetail.Text = TekstInfo(1).Text
    KwijtingBoeken!PolisDetail.Col = 5
    If KwijtingBoeken!PolisDetail.Text = "" Then
        KwijtingBoeken!PolisDetail.Text = "TODO: XML/TB2 retourblok importeren"
    End If
    KwijtingBoeken!PolisDetail.Col = 6
    KwijtingBoeken!PolisDetail.Text = TekstInfo(5).Text
            
    With KwijtingBoeken!PolisDetail
        '.SelStartCol = 0
        '.SelEndCol = 4
        '.SelStartRow = .Row
        '.SelEndRow = .Row
    End With
    
    GridText = "OK"
    Unload Me
End If

End Sub

Private Sub TekstInfo_GotFocus(Index As Integer)

TekstInfo(Index).SelLength = Len(TekstInfo(Index).Text)

End Sub

Private Sub TekstInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

On Local Error Resume Next
If KeyCode = 13 Then
    SendKeys "{TAB}", True
    Exit Sub
End If

If Index = 1 Then
    If KeyCode = 17 Then
        SharedFl = TABLE_CUSTOMERS
        aIndex = 1
        GridText = TekstInfo(1).Text
        SqlSearch.Show 1
        If Ktrl Then
            TekstInfo(1).Text = "-"
            KwijtingEdit!Ok.Enabled = False
            Exit Sub
        Else
            RecordToVeld TABLE_CUSTOMERS
            TekstInfo(1).Text = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
            bGet TABLE_CONTRACTS, 1, vBibTekst(TABLE_CUSTOMERS, "#A110 #")
            If Ktrl Or vSet(KEY_BUF(TABLE_CONTRACTS), 12) <> vSet(vBibTekst(TABLE_CUSTOMERS, "#A110 #"), 12) Then
                MsgBox "Geen polissen voor deze klant te vinden !!"
                TekstInfo(1).Text = "-"
                KwijtingEdit!Ok.Enabled = False
                Exit Sub
            Else
                Do
                    RecordToVeld TABLE_CONTRACTS
                    If Left(KwijtingBoeken!KeuzeInfo(0).Text, 4) = vBibTekst(TABLE_CONTRACTS, "#A010 #") Then
                        Msg = "Boeking voor polisnummer : " + vBibTekst(TABLE_CONTRACTS, "#A000 #") + vbCrLf + vbCrLf
                        Msg = Msg + vBibTekst(TABLE_CONTRACTS, "#vs99 #") + vbCrLf
                        Msg = Msg + vBibTekst(TABLE_CONTRACTS, "#vs98 #") + vbCrLf + vbCrLf
                        Msg = Msg + "Bent U zeker ?"
                        KtrlBox = MsgBox(Msg, vbYesNo + vbDefaultButton2)
                        If KtrlBox = vbYes Then
                            TekstInfo(0).Text = vBibTekst(TABLE_CONTRACTS, "#A000 #")
                            TekstInfo(0).SetFocus
                            Exit Do
                        End If
                    End If
                    bNext TABLE_CONTRACTS
                    If Ktrl Or vSet(KEY_BUF(TABLE_CONTRACTS), 12) <> vSet(vBibTekst(TABLE_CUSTOMERS, "#A110 #"), 12) Then
                        MsgBox "Geen polissen meer voor deze klant te vinden !!"
                        TekstInfo(1).Text = "-"
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
        bGet TABLE_CONTRACTS, 0, vSet((TekstInfo(0).Text), 12)
        If Ktrl Then
            KwijtingEdit!Ok.Enabled = False
            TekstInfo(1).SetFocus
            TekstInfo(0).Text = ""
            Exit Sub
        Else
            RecordToVeld TABLE_CONTRACTS
            If Left(KwijtingBoeken!KeuzeInfo(0).Text, 4) = vBibTekst(TABLE_CONTRACTS, "#A010 #") Then
                TekstInfo(0).Text = vBibTekst(TABLE_CONTRACTS, "#A000 #")
                bGet TABLE_CUSTOMERS, 0, vBibTekst(TABLE_CONTRACTS, "#A110 #")
                If Ktrl Then
                    TekstInfo(1).Text = "KlantLink onmogelijk !!! Kontroleer !!!"
                    KwijtingEdit!Ok.Enabled = False
                Else
                    RecordToVeld TABLE_CUSTOMERS
                    TekstInfo(1).Text = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
                    If Left(KwijtingBoeken!KeuzeInfo(1).Text, 1) = "1" Then
                        'TekstInfo(2).Text = Format(Val(vBibTekst(TABLE_CONTRACTS, "#v165 #")), "00") + "/" + Format(Val(vBibTekst(TABLE_CONTRACTS, "#v164 #")), "00") + "/" + Right(KwijtingBoeken!TekstInfo(0).Text, 4)
                        TekstInfo(2).Text = Format(Val(vBibTekst(TABLE_CONTRACTS, "#v165 #")), "00") + "/" + Right(KwijtingBoeken!TekstInfo(0).Text, 7)
                    Else
                        TekstInfo(2).Text = KwijtingBoeken!TekstInfo(0).Text 'Format(Val(vBibTekst(TABLE_CONTRACTS, "#v165 #")), "00") + "/" + Format(Val(vBibTekst(TABLE_CONTRACTS, "#v164 #")), "00") + "/" + Right(KwijtingBoeken!TekstInfo(0).Text, 4)
                    End If
                    Ok.Enabled = True
                End If
            Else
                MsgBox "Maatschappij : " + Left(KwijtingBoeken!KeuzeInfo(0).Text, 4) + " <> " + vBibTekst(TABLE_CONTRACTS, "#A010 #")
                TekstInfo(Index).Text = ""
            End If
        End If

    Case 2
        If DATE_INVALID((TekstInfo(2).Text)) Then
            TekstInfo(2).Text = KwijtingBoeken!TekstInfo(0).Text
            Beep
            TekstInfo(2).SetFocus
        End If
    Case 3, 4
        If bhEuro Then
            TekstInfo(Index).Text = Dec$(Val(TekstInfo(Index).Text), MASK_EUR)
        Else
            TekstInfo(Index).Text = Dec$(Val(TekstInfo(Index).Text), MASK_BEF)
        End If
End Select

End Sub

