VERSION 5.00
Begin VB.Form BeheerKwijtingen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Beheer Kwijtingen"
   ClientHeight    =   2655
   ClientLeft      =   1950
   ClientTop       =   1845
   ClientWidth     =   6765
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
   ScaleHeight     =   2655
   ScaleWidth      =   6765
   Begin VB.ComboBox CmbReedsBewerkt 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   150
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   1440
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6240
      Top             =   900
   End
   Begin VB.CommandButton Ok 
      Caption         =   "OK"
      Enabled         =   0   'False
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
      Left            =   5370
      TabIndex        =   6
      Top             =   30
      Width           =   1240
   End
   Begin VB.CommandButton Annuleren 
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
      Left            =   5370
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   1240
   End
   Begin VB.ComboBox Verzoek 
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   0
      Width           =   5055
   End
   Begin VB.ComboBox Operatiekode 
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
      Index           =   1
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   330
      Width           =   5055
   End
   Begin VB.TextBox VeldTekstVariabel 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   2220
      TabIndex        =   0
      Top             =   1350
      Width           =   4095
   End
   Begin VB.TextBox VeldTekstVariabel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   1410
      TabIndex        =   2
      Top             =   1950
      Width           =   1335
   End
   Begin VB.TextBox VeldTekstVariabel 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   1410
      TabIndex        =   4
      Top             =   2250
      Width           =   1335
   End
   Begin VB.TextBox VeldTekstVariabel 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   4500
      TabIndex        =   1
      Top             =   1650
      Width           =   1815
   End
   Begin VB.TextBox VeldTekstVariabel 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   4500
      TabIndex        =   3
      Top             =   1950
      Width           =   1815
   End
   Begin VB.TextBox VeldTekstVariabel 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   4500
      TabIndex        =   5
      Top             =   2250
      Width           =   1815
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Totaal premie"
      Height          =   240
      Index           =   5
      Left            =   30
      TabIndex        =   18
      Top             =   1950
      Width           =   1335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Klant"
      Height          =   285
      Index           =   3
      Left            =   1590
      TabIndex        =   16
      Top             =   1350
      Width           =   600
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Commissie"
      Height          =   285
      Index           =   6
      Left            =   390
      TabIndex        =   19
      Top             =   2220
      Width           =   975
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Maatschappij"
      Height          =   270
      Index           =   2
      Left            =   1590
      TabIndex        =   12
      Top             =   720
      Width           =   1440
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Polis"
      Height          =   270
      Index           =   0
      Left            =   4350
      TabIndex        =   10
      Top             =   720
      Width           =   600
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Producent"
      Height          =   270
      Index           =   1
      Left            =   150
      TabIndex        =   11
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Onze Referte"
      Height          =   285
      Index           =   4
      Left            =   2850
      TabIndex        =   17
      Top             =   1650
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Aanvangsdatum"
      Height          =   285
      Index           =   7
      Left            =   2850
      TabIndex        =   20
      Top             =   1950
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nummer Kwijting"
      Height          =   285
      Index           =   8
      Left            =   2850
      TabIndex        =   21
      Top             =   2250
      Width           =   1575
   End
   Begin VB.Label VeldTekstVast 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   2
      Left            =   4320
      TabIndex        =   15
      Top             =   990
      Width           =   1800
   End
   Begin VB.Label VeldTekstVast 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   13
      Top             =   990
      Width           =   1320
   End
   Begin VB.Label VeldTekstVast 
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Index           =   1
      Left            =   1560
      TabIndex        =   14
      Top             =   990
      Width           =   2640
   End
End
Attribute VB_Name = "BeheerKwijtingen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim A As String * 140
Dim aa As String * 170

Private Sub Annuleren_Click()

Unload Me

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Verzoek.AddItem "027: B.R.B. (terugkeer kwijting naar maatschappij)"
Verzoek.AddItem "028: A.T.K. (aanvraag teruggave kwijting)"
Verzoek.ListIndex = 0
Timer1.Enabled = True

Mim.Arrange 1

End Sub


Private Sub Ok_Click()

If VeldTekstVast(0).Caption + VeldTekstVast(1).Caption + VeldTekstVast(2).Caption = "" Then
    Beep
    Exit Sub
Else
    For Tel = 0 To CmbReedsBewerkt.ListCount - 1
        CmbReedsBewerkt.ListIndex = Tel
        If vBibTekst(TABLE_CONTRACTS, "#A000 #") = CmbReedsBewerkt.Text Then
            MsgBox "U hebt in deze sessie reeds een BRB/ATK doorgegeven"
            Exit Sub
        End If
    Next
    aa = ""
    Mid(aa, 1, 12) = VeldTekstVast(2).Caption
    Mid(aa, 13, 8) = VeldTekstVast(0).Caption
    Mid(aa, 21, 35) = vSet((VeldTekstVariabel(0).Text), 35)
    Mid(aa, 56, 10) = vSet((VeldTekstVariabel(1).Text), 10)
    Mid(aa, 71, 9) = Tk((VeldTekstVariabel(2).Text), 9)
    Mid(aa, 80, 8) = Tk((VeldTekstVariabel(3).Text), 8)
    Mid(aa, 88, 8) = Puur(Left(VeldTekstVariabel(4).Text, 10))
    Mid(aa, 96, 9) = Tk((VeldTekstVariabel(5).Text), 9)
    Mid(aa, 105, 2) = Left(Operatiekode(1).Text, 2)
    Msg$ = Verzoek.Text + vbCrLf + vbCrLf + Operatiekode(1).Text + vbCrLf + vbCrLf + "Wegschrijven naar " + LOCATION_ASWEB + "as1.snd" + vbCrLf + vbCrLf + " Bent U zeker ?"
    KtrlBox = MsgBox(Msg$, vbYesNo, "B.R.B. / A.T.K.")
    If KtrlBox = vbYes Then
        VoorVoegsel
        SchrijfWeg (A + aa)
        VeldTekstVariabel(0).SetFocus
        CmbReedsBewerkt.AddItem vBibTekst(TABLE_CONTRACTS, "#A000 #")
        Unload Me
    End If
End If

End Sub

Private Sub OperatieKeuze()

Operatiekode(1).Clear
If Left(Verzoek.Text, 3) = "028" Then
    Operatiekode(1).AddItem "21: Kwijting naar Producent, betaald aan Producent"
    Operatiekode(1).AddItem "22: Kwijting naar kliënt, betaald aan Producent"
    Operatiekode(1).AddItem "31: Deze + volgende kwijtingen naar Producent"
    Operatiekode(1).AddItem "32: Deze + volgende, inning producent, kwijting naar kliënt"
    Operatiekode(1).AddItem "41: Enkel deze kwijting naar/inning door Producent"
    Operatiekode(1).AddItem "42: Enkel deze kwijting inning producent, sturen naar kliënt"
    Operatiekode(1).ListIndex = 0
Else
    Operatiekode(1).AddItem "B : Wijziging of bijvoegsel werd aangevraagd"
    Operatiekode(1).AddItem "G : Weigering automatische indexering (Leven)"
    Operatiekode(1).AddItem "S : Inning van deze kwijting door de maatschappij"
    Operatiekode(1).AddItem "T : Inning deze en volgende door de maatschappij"
    Operatiekode(1).ListIndex = 2
End If
Exit Sub

End Sub


Private Sub SchrijfWeg(RecordString As String)
Dim FlNetTemp As Integer
On Local Error GoTo SchrijfWegError

FlNetTemp = FreeFile
Open LOCATION_ASWEB + "AS1.SND" For Append Lock Write As FlNetTemp
Print #FlNetTemp, RecordString
Close FlNetTemp
    Exit Sub

SchrijfWegError:
    MsgBox "Openen van " + LOCATION_ASWEB + "AS1.SND onmogelijk.  Probeer opnieuw"
    Resume

End Sub


Private Sub Timer1_Timer()

If RTrim$(VeldTekstVast(2).Caption + Venster.TxtInfo(0).Text) = "" Then
    SnelHelpPrint "Aktiveer een polisnummer via POLISINFO a.u.b. ! [CTRL + P]", BL_LOGGING
    Ok.Enabled = False
ElseIf Venster.TxtInfo(0).Text <> VeldTekstVast(2).Caption Then
    VeldTekstVast(2).Caption = Venster.TxtInfo(0).Text
End If

End Sub

Private Sub VeldTekstVariabel_GotFocus(Index As Integer)

VeldTekstVariabel(Index).SelLength = Len(VeldTekstVariabel(Index).Text)

End Sub

'
Private Sub VeldTekstVariabel_LostFocus(Index As Integer)

Select Case Index
    Case 4
        If DATE_INVALID(Left(VeldTekstVariabel(4).Text, 10)) Then
            Beep
            VeldTekstVariabel(4).Text = MIM_GLOBAL_DATE
            VeldTekstVariabel(4).SetFocus
        End If
End Select

End Sub

Private Sub VeldTekstVast_Change(Index As Integer)

If Index <> 2 Then Exit Sub

bGet TABLE_CONTRACTS, 0, VeldTekstVast(2).Caption
If Ktrl Then
    Ok.Enabled = False
Else
    RecordToVeld TABLE_CONTRACTS
    'If KTRLVerzoek(Left(Verzoek.Text, 3), vBibTekst(TABLE_CONTRACTS, "#A010 #")) Then
    'Else
    '    MsgBox "Dit verzoek is voor deze maatschappij niet mogelijk.  Gebruik vrije berichtgeving als alternatief !", , "Beheer kwijtingen"
    '    Ok.Enabled = False
    '    Exit Sub
    'End If
    bGet TABLE_SUPPLIERS, 0, "CO" + vBibTekst(TABLE_CONTRACTS, "#A010 #")
    If Ktrl Then
        SnelHelpPrint "Aktiveer een polisnummer via POLISINFO a.u.b. ! [CTRL + P]", BL_LOGGING
    Else
        RecordToVeld TABLE_SUPPLIERS
        VeldTekstVast(0).Caption = vBibTekst(TABLE_SUPPLIERS, "#A400 #")
        VeldTekstVast(1).Caption = vBibTekst(TABLE_SUPPLIERS, "#A100 #")
        VeldTekstVariabel(0).Text = Venster!TxtInfo(1).Text + " " + Venster!TxtInfo(3).Text
        VeldTekstVariabel(1).Text = Venster!TxtInfo(18).Text
        VeldTekstVariabel(2).Text = vBibTekst(TABLE_CONTRACTS, "#B010 #")
        VeldTekstVariabel(3).Text = vBibTekst(TABLE_CONTRACTS, "#B014 #")
        VeldTekstVariabel(4).Text = vBibTekst(TABLE_CONTRACTS, "#v165 #") + "/" + vBibTekst(TABLE_CONTRACTS, "#v164 #") + "/" + Left(PERIOD_FROMTO, 4)
        VeldTekstVariabel(5).Text = ""
        Ok.Enabled = True
    End If
End If

End Sub

Private Sub Verzoek_Click()

OperatieKeuze

End Sub

Private Sub VoorVoegsel()

A = ""
Mid(A, 1, 3) = "AS1"                           'constante
Mid(A, 16, 1) = "0"                            'constante
Mid(A, 17, 3) = Left(Verzoek.Text, 3)
Mid(A, 20, 1) = "1"                            'gestruktureerd
Mid(A, 21, 1) = "1"                            'producent
Mid(A, 22, 8) = vSet(ProducentNummer, 8)       'ASNET prod.nr.
Mid(A, 30, 12) = Format(Now, "yymmddhhmmss")  'datum
Mid(A, 42, 1) = "1"                            'gewoon niet vertrouwelijk
Mid(A, 43, 1) = "2"                            'bestemmeling=mij.
Mid(A, 44, 4) = vBibTekst(TABLE_CONTRACTS, "#A010 #")  'BVVO kode
Mid(A, 48, 2) = "01"                   'volgnummer
Mid(A, 50, 1) = "X"                    'einde tekst

End Sub

