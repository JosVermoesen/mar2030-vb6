VERSION 5.00
Begin VB.Form WijzigenInlichtingen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wijzigingen Inlichtingen "
   ClientHeight    =   2370
   ClientLeft      =   270
   ClientTop       =   3345
   ClientWidth     =   8880
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
   ScaleHeight     =   2370
   ScaleWidth      =   8880
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3060
      Top             =   900
   End
   Begin VB.CommandButton Inlichtingen 
      Caption         =   "&Inlichtingen"
      Enabled         =   0   'False
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1680
   End
   Begin VB.CommandButton Sluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   360
      Left            =   1920
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1560
   End
   Begin VB.TextBox TekstInfo 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   1
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2640
   End
   Begin VB.ComboBox KeuzeInfo 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   4320
   End
   Begin VB.ComboBox KeuzeInfo 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   4320
   End
   Begin VB.CheckBox CheckBox 
      Caption         =   "Korrespondentieadres niet schrappen"
      Height          =   240
      Index           =   2
      Left            =   4740
      TabIndex        =   3
      Top             =   480
      Width           =   3780
   End
   Begin VB.CheckBox CheckBox 
      Caption         =   "Geen naamverbetering"
      Height          =   240
      Index           =   1
      Left            =   4740
      TabIndex        =   2
      Top             =   240
      Width           =   2265
   End
   Begin VB.CheckBox CheckBox 
      Caption         =   "Alle polissen klient"
      Height          =   240
      Index           =   0
      Left            =   4740
      TabIndex        =   1
      Top             =   0
      Width           =   1965
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Polisnummer"
      Height          =   255
      Left            =   210
      TabIndex        =   9
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Klantinfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   1350
      Left            =   3810
      TabIndex        =   8
      Top             =   810
      Width           =   5040
   End
End
Attribute VB_Name = "WijzigenInlichtingen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim A As String * 140
Dim aa As String * 413
Dim FlAsNet As Integer






Private Sub Form_Load()

InstallMij
KeuzeInfo(1).AddItem "022: Inlichting Klient"
KeuzeInfo(1).AddItem "024: Inlichting Polis Auto"
KeuzeInfo(1).ListIndex = 0
Timer1.Enabled = True
Mim.Arrange 1

End Sub



Private Sub Form_Unload(Cancel As Integer)

Timer1.Enabled = False

End Sub

Private Sub Inlichtingen_Click()
Dim PolisNummer     As String * 12
Dim MijNaam         As String
Dim TypeBericht     As String
Dim TeleBibTekst    As String
Dim Msg             As String
Dim OptieTxt        As String
Dim TempoTxt        As String
Dim CrText          As String
Dim CrText2         As String
Dim fKtrl           As Integer
Dim T               As Integer
Dim DataPos         As Integer
Dim DataLen         As Integer

Timer1.Enabled = False
PolisNummer = LTrim$(RTrim$(TekstInfo(1).Text))
Select Case Left(KeuzeInfo(1).Text, 3)
    Case "022"
        TypeBericht = "Wijzigen Inlichtingen Klient "
    Case "024"
        TypeBericht = "Wijzigen Polis AUTO "
    Case Else
        MsgBox "STOP"
End Select

bGet TABLE_VARIOUS, 1, "25" + Left(KeuzeInfo(0).Text, 4) + PolisNummer
If Ktrl Then
    fKtrl = 99
    TLB_RECORD(TABLE_VARIOUS) = ""
    vBib TABLE_VARIOUS, PolisNummer, "A000"
    vBib TABLE_VARIOUS, Right(KeuzeInfo(1).Text, 8), "A400"
Else
    fKtrl = 100
    RecordToVeld TABLE_VARIOUS
End If

Msg = "Klantinformatie automatisch inladen voor deze polis. Bent U zeker ?"
KtrlBox = MsgBox(Msg, vbYesNo)
If KtrlBox = vbYes Then
    If TLBPag3("AS1" + Left(KeuzeInfo(1).Text, 3)) Then
    Else
        MsgBox "Stop"
    End If
    T = 0
    Do While TELEBIB_CODE(T) <> Space$(10)
        If RTrim$(vBibTekst(TABLE_CUSTOMERS, "#" + Mid(TELEBIB_CODE(T), 5, 5) + "#")) <> "" Then
            vBib TABLE_VARIOUS, vBibTekst(TABLE_CUSTOMERS, "#" + Mid(TELEBIB_CODE(T), 5, 5) + "#"), Mid(TELEBIB_CODE(T), 5, 5)
        End If
        T = T + 1
    Loop
End If
If TeleBibClick(1000 + Val(Left(KeuzeInfo(1).Text, 3))) Then
    Msg = "Gestructureerd verzoek wijziging inlichtingen wegschrijven, bent U zeker ?" + vbCrLf
    KtrlBox = MsgBox(Msg, vbYesNo, (KeuzeInfo(1).Text))
    If KtrlBox = vbYes Then
        vBib TABLE_VARIOUS, Left(KeuzeInfo(0).Text, 4), "A010"
        vBib TABLE_VARIOUS, PolisNummer, "A000"
        vBib TABLE_VARIOUS, "K" + vBibTekst(TABLE_CUSTOMERS, "#A110 #"), "v004"
        vBib TABLE_VARIOUS, "25" + Left(KeuzeInfo(0).Text, 4) + PolisNummer, "v005"
        If fKtrl = 99 Then
            bInsert TABLE_VARIOUS, 1
            If Ktrl Then MsgBox "stop"
        Else
            bUpdate TABLE_VARIOUS, 1
            If Ktrl Then MsgBox "stop"
        End If
        VoorVoegsel
        T = 0
        Do While TELEBIB_CODE(T) <> Space$(10)
            TempoTxt = Space$(TELEBIB_LENGTH(T))
            Mid(TempoTxt, 1) = vBibTekst(TABLE_VARIOUS, "#" + Mid(TELEBIB_CODE(T), 5, 5) + "#")
            Mid(aa, TELEBIB_POS(T)) = TempoTxt
            T = T + 1
        Loop
        Mid(aa, 1, 12) = TekstInfo(1).Text
        Mid(aa, 13, 8) = Right(KeuzeInfo(0).Text, 8)
        Select Case Left(KeuzeInfo(1).Text, 3)
            Case "022"
                Mid(aa, 33, 1) = Format(CheckBox(0).Value + 1, "0")
                Mid(aa, 34, 1) = Format(CheckBox(1).Value + 1, "0")
                Mid(aa, 35, 1) = Format(CheckBox(2).Value + 1, "0")
        End Select
        SchrijfWeg A + aa
        Inlichtingen.Enabled = False
        Unload Me
    End If
End If
Timer1.Enabled = True
    
End Sub

Private Sub InstallMij()
Dim T As Integer

KeuzeInfo(0).Clear
T = -1
bGetOrGreater TABLE_SUPPLIERS, 0, vSet("CO", 12)
If Ktrl Then
    Beep
    Exit Sub
Else
    RecordToVeld TABLE_SUPPLIERS
End If
Screen.MousePointer = vbHourglass
If Left(KEY_BUF(TABLE_SUPPLIERS), 2) <> "CO" Then
    Beep
Else
    T = T + 1
    KeuzeInfo(0).AddItem Mid(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 3, 4) + ": " + vBibTekst(TABLE_SUPPLIERS, "#A100 #") + "/" + vSet(vBibTekst(TABLE_SUPPLIERS, "#A400 #"), 8)
    Do
        bNext TABLE_SUPPLIERS
        If Ktrl Or Left(KEY_BUF(TABLE_SUPPLIERS), 2) <> "CO" Then
            Exit Do
        Else
            RecordToVeld TABLE_SUPPLIERS
            T = T + 1
            KeuzeInfo(0).AddItem Mid(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 3, 4) + ": " + vBibTekst(TABLE_SUPPLIERS, "#A100 #") + "/" + vSet(vBibTekst(TABLE_SUPPLIERS, "#A400 #"), 8)
        End If
    Loop
End If
Screen.MousePointer = vbNormal
If KeuzeInfo(0).ListCount Then
    KeuzeInfo(0).ListIndex = 0
End If

End Sub

Private Sub KeuzeInfo_Click(Index As Integer)

Select Case Index
    Case 1
        Select Case KeuzeInfo(1).ListIndex
            Case 0
                For Tel = 0 To 2
                    CheckBox(Tel).Visible = True
                    CheckBox(Tel).Value = 0
                Next
                CheckBox(2).Value = 1

            Case 1
                For Tel = 0 To 2
                    CheckBox(Tel).Visible = False
                Next
                TekstInfo_Change 1
            Case Else
                MsgBox "Stop"
        End Select
End Select

End Sub


Private Sub Schoon()

KeuzeInfo(1).ListIndex = 0
TekstInfo(1).Text = ""

End Sub

Private Sub SchrijfWeg(RecordLijn As String)

On Local Error GoTo SchrijfWegError

Print #FlAsNet, RecordLijn
Close FlAsNet
Exit Sub

SchrijfWegError:
    FlAsNet = FreeFile
    Open LOCATION_ASWEB + "AS1.SND" For Append Lock Write As FlAsNet
    Resume

End Sub

Private Sub sluiten_Click()

Unload Me

End Sub


Private Sub VoorVoegsel()

A = ""
aa = ""
Mid(A, 1, 3) = "AS1"                               'constante
Mid(A, 16, 1) = "0"                                'constante
Mid(A, 17, 3) = Left(KeuzeInfo(1).Text, 3)        'Klient of auto
Mid(A, 20, 1) = "1"                                'gestruktureerde inhoud
Mid(A, 21, 1) = "1"                                'producent
Mid(A, 22, 8) = vSet(ProducentNummer, 8)        'ASNET prod.nr.
Mid(A, 30, 12) = Format(Now, "yymmddhhmmss")      'datum
Mid(A, 42, 1) = "1"                                'gewoon niet vertrouwelijk
Mid(A, 43, 1) = "2"                                'bestemmeling=mij.
Mid(A, 44, 4) = Left(KeuzeInfo(0).Text, 4)        'BVVO kode
Mid(A, 48, 2) = "01"                               'volgnummer
Mid(A, 50, 1) = "X"                                'vervolg tekst
'Mid(a, 91, 1) = "P"                                'Referentietype
'Mid(a, 92, 12) = VeldSet((TekstInfo(1).Text), 12)  'Referentie nr
'Mid(a, 104, 1) = "A"                               'dienstkode

End Sub


Private Sub TekstInfo_Change(Index As Integer)

bGet TABLE_CONTRACTS, 0, TekstInfo(1).Text
If Ktrl Then
    Inlichtingen.Enabled = False
Else
    RecordToVeld TABLE_CONTRACTS
    mijflag = False
    For Tel = 0 To KeuzeInfo(0).ListCount - 1
        If vBibTekst(TABLE_CONTRACTS, "#A010 #") = Left(KeuzeInfo(0).List(Tel), 4) Then
            KeuzeInfo(0).ListIndex = Tel
            If KTRLVerzoek(Left(KeuzeInfo(1).Text, 3), Left(KeuzeInfo(0).Text, 4)) Then
                mijflag = True
            Else
                MsgBox "Dit verzoek is voor deze maatschappij niet mogelijk.  Gebruik vrije berichtgeving als alternatief !"
            End If
         End If
    Next
    If mijflag = False Then
        Inlichtingen.Enabled = False
    Else
        KlantInfo.Caption = Venster!TxtInfo(18).Text + vbCrLf + Venster!TxtInfo(1).Text + vbCrLf + Venster!TxtInfo(3).Text
        Inlichtingen.Enabled = True
    End If
End If

End Sub

Private Sub Timer1_Timer()

If RTrim$(TekstInfo(1).Text + Venster.TxtInfo(0).Text) = "" Then
    SnelHelpPrint "Aktiveer een polisnummer via POLISINFO a.u.b. ! [CTRL + P]", BL_LOGGING
    Inlichtingen.Enabled = False
ElseIf Venster.TxtInfo(0).Text <> TekstInfo(1).Text Then
    TekstInfo(1).Text = Venster.TxtInfo(0).Text
    'Inlichtingen.Enabled = False
    SnelHelpPrint "", BL_LOGGING
End If

End Sub


