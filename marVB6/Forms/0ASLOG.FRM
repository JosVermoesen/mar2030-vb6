VERSION 5.00
Begin VB.Form LogBoekKontrole 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detail van het verzoek"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   8895
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
   ScaleHeight     =   6360
   ScaleWidth      =   8895
   Begin VB.CommandButton CmdAfdrukken 
      Caption         =   "Af&druk"
      Height          =   360
      Left            =   7620
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1200
   End
   Begin VB.TextBox VrijeTekst 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3675
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2100
      Width           =   8685
   End
   Begin VB.CommandButton Sluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   360
      Left            =   6300
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5880
      Width           =   1200
   End
   Begin VB.CommandButton Verder 
      Caption         =   ">>"
      Height          =   360
      Left            =   5550
      TabIndex        =   2
      Top             =   5880
      Width           =   690
   End
   Begin VB.CommandButton VanafBegin 
      Caption         =   "Vanaf &Begin"
      Height          =   360
      Left            =   3660
      TabIndex        =   1
      Top             =   5880
      Width           =   1800
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "----.----1----.----2----.----3----.----4----.----5----.----6----.----7----.----"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   25
      Top             =   1890
      Width           =   8385
   End
   Begin VB.Label LabelInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   1170
      TabIndex        =   11
      Top             =   570
      Width           =   2655
   End
   Begin VB.Label LabelInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   1170
      TabIndex        =   12
      Top             =   810
      Width           =   2385
   End
   Begin VB.Label LabelInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   8
      Left            =   2940
      TabIndex        =   14
      Top             =   1050
      Width           =   1935
   End
   Begin VB.Label LabelInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   7
      Left            =   1170
      TabIndex        =   15
      Top             =   1050
      Width           =   1755
   End
   Begin VB.Label LabelInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   9
      Left            =   1170
      TabIndex        =   18
      Top             =   1290
      Width           =   1440
   End
   Begin VB.Label LabelInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   10
      Left            =   1170
      TabIndex        =   17
      Top             =   1530
      Width           =   4335
   End
   Begin VB.Label LabelInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   1170
      TabIndex        =   20
      Top             =   90
      Width           =   4215
   End
   Begin VB.Label LabelInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   5250
      TabIndex        =   16
      Top             =   330
      Width           =   1485
   End
   Begin VB.Label LabelInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   5250
      TabIndex        =   22
      Top             =   570
      Width           =   1155
   End
   Begin VB.Label LabelInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1170
      TabIndex        =   19
      Top             =   330
      Width           =   2655
   End
   Begin VB.Label LabelInfo 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   3540
      TabIndex        =   13
      Top             =   810
      Width           =   1875
   End
   Begin VB.Label LabelInfo 
      BorderStyle     =   1  'Fixed Single
      Height          =   270
      Index           =   11
      Left            =   1320
      TabIndex        =   24
      Top             =   5940
      Width           =   2295
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Betreft"
      Height          =   240
      Index           =   10
      Left            =   210
      TabIndex        =   7
      Top             =   1530
      Width           =   840
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Beheer"
      Height          =   240
      Index           =   12
      Left            =   210
      TabIndex        =   9
      Top             =   1290
      Width           =   840
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Zetel"
      Height          =   240
      Index           =   13
      Left            =   4530
      TabIndex        =   10
      Top             =   330
      Width           =   600
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datum"
      Height          =   240
      Index           =   0
      Left            =   4410
      TabIndex        =   4
      Top             =   570
      Width           =   720
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Referte"
      Height          =   240
      Index           =   11
      Left            =   210
      TabIndex        =   8
      Top             =   1050
      Width           =   840
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tak"
      Height          =   240
      Index           =   7
      Left            =   210
      TabIndex        =   6
      Top             =   810
      Width           =   840
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dienst"
      Height          =   240
      Index           =   6
      Left            =   210
      TabIndex        =   5
      Top             =   570
      Width           =   840
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Verzoek"
      Height          =   240
      Index           =   14
      Left            =   120
      TabIndex        =   21
      Top             =   60
      Width           =   960
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Assurnet IDA"
      Height          =   255
      Index           =   16
      Left            =   90
      TabIndex        =   23
      Top             =   5940
      Width           =   1185
   End
End
Attribute VB_Name = "LogBoekKontrole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim AS1Mij(90) As String
Dim AS1Dienst(5) As String
Dim AS1Tak(9) As String
Dim AS1Klient(2) As String
Dim AS1RefType(2) As String
Dim AS1Zetel(19) As String

Dim FlTemp As Integer
Dim aa As String

Dim FunktieTekst(99) As String
Dim StruktuurTekst(1) As String
Dim TakTekst(9) As String

Dim IDA As String * 12

Private Sub cmdAfdrukken_Click()

'MsgBox "Grafische schermafdruk is tijdelijk niet mogelijk"
On Error Resume Next
PrintForm
If Err Then MsgBox Error

End Sub

Private Sub Form_Load()

InitASBestanden
InstallMij

AS1Dienst(0) = "A: Produktie/uitgifte"
AS1Dienst(1) = "B: Boekhouding"
AS1Dienst(2) = "C: Schadegeval"
AS1Dienst(3) = "D: Commercieel"
AS1Dienst(4) = "E: Informatica en Organisatie"
AS1Dienst(5) = "H: Herverzekering"

AS1Tak(0) = "1: Leven"
AS1Tak(1) = "2: Ziekte en individuele"
AS1Tak(2) = "3: Brand - gewone risiko's"
AS1Tak(3) = "4: Burgerlijke aansprakelijkheid"
AS1Tak(4) = "5: Auto"
AS1Tak(5) = "6: Arbeidsongevallen"
AS1Tak(6) = "7: Brand - industriële risiko's"
AS1Tak(7) = "8: Transport"
AS1Tak(8) = "9: Diverse"
AS1Tak(9) = "0: Leningen"

AS1Klient(0) = "O: Ondernemingen"
AS1Klient(1) = "P: Partikulieren"
AS1Klient(2) = " : Onbepaald "

AS1RefType(0) = "P: Polis"
AS1RefType(1) = "S: Schadegeval"
AS1RefType(2) = " : Andere"

AS1Zetel(0) = "10: Brussel"
AS1Zetel(1) = "20: Antwerpen"
AS1Zetel(2) = "23: Turnhout"
AS1Zetel(3) = "27: Sint-Niklaas"
AS1Zetel(4) = "28: Mechelen"
AS1Zetel(5) = "30: Leuven"
AS1Zetel(6) = "35: Hasselt"
AS1Zetel(7) = "40: Liège"
AS1Zetel(8) = "48: Verviers"
AS1Zetel(9) = "49: Eupen"
AS1Zetel(10) = "50: Namur"
AS1Zetel(11) = "60: Charleroi"
AS1Zetel(12) = "67: Arlon"
AS1Zetel(13) = "70: Mons"
AS1Zetel(14) = "75: Tournai"
AS1Zetel(15) = "80: Brugge"
AS1Zetel(16) = "84: Oostende"
AS1Zetel(17) = "85: Kortrijk"
AS1Zetel(18) = "90: Gent"
AS1Zetel(19) = "93: Aalst"

Screen.MousePointer = vbHourglass
If XLogKey <> "" Then
    IDA = Mid(XLogKey, 1, 3) + Mid(XLogKey, 5, 5) + Mid(XLogKey, 11, 4)
    If LaadInOK() Then
        LabelsVullen
    End If
    If LabelInfo(0).Caption = "" Then
        MsgBox "Gestruktureerde informatie voor " + vbCrLf + vbCrLf + XLogKey + vbCrLf + vbCrLf + "is reeds verwijderd."
        Verder.Enabled = False
        VanafBegin.Default = True
    End If
End If
Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Unload(Cancel As Integer)

Close FlTemp
FlTemp = 0

End Sub

Private Sub InitASBestanden()
Dim T As String
Dim TT As Integer
Dim T1 As Integer

On Local Error GoTo InitHandler

T1 = FreeFile
Open PROGRAM_LOCATION + "Def\" + "as1oct.ini" For Input As T1
Line Input #T1, T
Line Input #T1, T
For TT = 0 To Val(T)
    Line Input #T1, StruktuurTekst(TT)
Next

Line Input #T1, T
Line Input #T1, T
For TT = 0 To Val(T)
    Line Input #T1, TakTekst(TT)
Next
Line Input #T1, T
Line Input #T1, T
For TT = 0 To Val(T)
    Line Input #T1, FunktieTekst(TT)
Next
Close T1
Exit Sub

InitHandler:
MsgBox "ASOCT.INI beschadigd of niet aanwezig.  Installer het programma opnieuw !"
End

End Sub

Private Sub InstallMij()
Dim T As Integer

T = -1
bGetOrGreater TABLE_SUPPLIERS, 0, "CO"
If Ktrl Then
    Exit Sub
Else
    RecordToVeld TABLE_SUPPLIERS
End If
Screen.MousePointer = vbHourglass
If Left(KEY_BUF(TABLE_SUPPLIERS), 2) <> "CO" Then
Else
    T = T + 1
    AS1Mij(0) = Mid(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 3, 4) + ": " + vBibTekst(TABLE_SUPPLIERS, "#A100 #") + "/" + vSet(vBibTekst(TABLE_SUPPLIERS, "#v014 #"), 8)
    Do
        bNext TABLE_SUPPLIERS
        If Ktrl Or Left(KEY_BUF(TABLE_SUPPLIERS), 2) <> "CO" Then
            Exit Do
        Else
            RecordToVeld TABLE_SUPPLIERS
            T = T + 1
            AS1Mij(T) = Mid(vBibTekst(TABLE_SUPPLIERS, "#A110 #"), 3, 4) + ": " + vBibTekst(TABLE_SUPPLIERS, "#A100 #") + "/" + vSet(vBibTekst(TABLE_SUPPLIERS, "#v014 #"), 8)
        End If
    Loop
End If
Screen.MousePointer = vbNormal

End Sub

Private Function LaadInOK()
Dim RecordLijn As String

LaadInOK = False
aa = ""

bGet TABLE_VARIOUS, 1, "26" + IDA
If Ktrl Then
    Exit Function
End If

Do
    RecordToVeld TABLE_VARIOUS
    GoSub Bijvoegen
    bNext TABLE_VARIOUS
    If vSet(KEY_BUF(TABLE_VARIOUS), 20) <> vSet("26" + IDA, 20) Then
        Stop
        Exit Do
    End If
Loop
Stop
If aa <> "" Then
    LaadInOK = True
End If
Exit Function

Bijvoegen:
RecordLijn = vBibTekst(TABLE_VARIOUS, "#v228 #")
If Mid(RecordLijn, 50, 1) = "X" Then
    If aa = "" Then
        aa = RecordLijn
    Else
        aa = aa + Right(RecordLijn, Len(RecordLijn) - 140)
    End If
    LaadInOK = True
    Exit Function
Else
    If aa = "" Then
        aa = RecordLijn
    Else
        aa = aa + Right(RecordLijn, Len(RecordLijn) - 140)
    End If
End If
Return

End Function

Private Sub LabelsVullen()
Dim T As Integer

For T = 0 To 11
    LabelInfo(0).Caption = ""
Next

If Val(Mid(aa, 17, 3)) > 88 Then
Else
    LabelInfo(0).Caption = FunktieTekst(Val(Mid(aa, 17, 3)))
End If

T = 0
Do While AS1Mij(T) <> ""
    If Mid(aa, 44, 4) = Left(AS1Mij(T), 4) Then
        LabelInfo(1).Caption = AS1Mij(T)
        Exit Do
    End If
    T = T + 1
Loop

For T = 0 To 19
    If Mid(aa, 117, 2) = Left(AS1Zetel(T), 2) Then
        LabelInfo(2).Caption = AS1Zetel(T)
        Exit For
    End If
Next

For T = 0 To 5
    If Mid(aa, 104, 1) = Left(AS1Dienst(T), 1) Then
        LabelInfo(3).Caption = AS1Dienst(T)
        Exit For
    End If
Next

LabelInfo(4).Caption = Mid(aa, 34, 2) + "/" + Mid(aa, 32, 2) + "/" + Mid(aa, 30, 2) + " " + Mid(aa, 36, 2) + ":" + Mid(aa, 38, 2) + ":" + Mid(aa, 40, 2)

For T = 0 To 9
    If Mid(aa, 105, 1) = Left(AS1Tak(T), 1) Then
        LabelInfo(5).Caption = AS1Tak(T)
        Exit For
    End If
Next

For T = 0 To 2
    If Mid(aa, 106, 1) = Left(AS1Klient(T), 1) Then
        LabelInfo(6).Caption = AS1Klient(T)
        Exit For
    End If
Next

For T = 0 To 2
    If Mid(aa, 91, 1) = Left(AS1RefType(T), 1) Then
        LabelInfo(7).Caption = AS1RefType(T)
        Exit For
    End If
Next

LabelInfo(8).Caption = Mid(aa, 92, 12)
LabelInfo(9).Caption = Mid(aa, 107, 10)
LabelInfo(10).Caption = Mid(aa, 51, 40)
LabelInfo(11).Caption = Mid(aa, 4, 4) + "/" + Mid(aa, 8, 5) + "/" + Mid(aa, 13, 3)

Select Case Mid(aa, 17, 3)
    Case "020"
        Msg$ = ""
        aa = Right(aa, Len(aa) - 140)
        Do While aa <> ""
            Msg$ = Msg$ + Left(aa, 79) + vbCrLf
            aa = Right(aa, Len(aa) - 79)
        Loop
        VrijeTekst.Text = Msg$
    Case "027", "028"
               Msg$ = "Polisnummer : " + Mid(aa, 140 + 1, 12) + vbCrLf
        Msg$ = Msg$ + "Producentnr : " + Mid(aa, 140 + 13, 8) + vbCrLf
        Msg$ = Msg$ + "Verzekerde  : " + Mid(aa, 140 + 21, 35) + vbCrLf
        Msg$ = Msg$ + "Referte     : " + Mid(aa, 140 + 56, 10) + vbCrLf + vbCrLf
        Msg$ = Msg$ + "Premie    : " + Format(Val(Mid(aa, 140 + 71, 9)), "#########") + vbCrLf
        Msg$ = Msg$ + "Commissie : " + Format(Val(Mid(aa, 140 + 80, 8)), "########") + vbCrLf + vbCrLf
        Msg$ = Msg$ + "Datum document : " + Mid(aa, 140 + 88, 8) + vbCrLf
        Msg$ = Msg$ + "Nr. kwijting   : " + Mid(aa, 140 + 96, 9) + vbCrLf
        Msg$ = Msg$ + "Operatie       : " + Mid(aa, 140 + 105, 2) + vbCrLf + vbCrLf
        Msg$ = Msg$ + "Opmerking : " + Mid(aa, 140 + 107, 50)
        VrijeTekst.Text = Msg$
    Case "912", "913"
        VrijeTekst.Text = aa
Case Else
    VrijeTekst.Text = aa
End Select

End Sub

Private Sub sluiten_Click()

Unload Me

End Sub

Private Sub VanafBegin_Click()

Close FlTemp
If LaadInOK() Then
    LabelsVullen
    Verder.Enabled = True
    Verder.Default = True
    VrijeTekst.SetFocus
Else
    'MsgBox "Het tekstbestand " + vbCrLf + vbCrLf + UCase$(LOCATION_COMPANYDATA) + "AS1SND.LOG" + vbCrLf + vbCrLf + "bestaat niet meer !"
End If

End Sub

Private Sub Verder_Click()

If LaadInOK() Then
    LabelsVullen
    VrijeTekst.SetFocus
Else
    Verder.Enabled = False
    VanafBegin.Default = True
End If

End Sub

Private Sub VrijeTekst_KeyPress(KeyAscii As Integer)

KeyAscii = 0

End Sub

