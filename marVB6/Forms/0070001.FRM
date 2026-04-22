VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form KwijtingBoeken 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kwijtingen boeken"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   1410
   ClientWidth     =   8760
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
   ScaleHeight     =   5310
   ScaleWidth      =   8760
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox cbTB2ktrl 
      Caption         =   "TB2 controle"
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   120
      Width           =   1455
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   9340
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Inboeking"
      TabPicture(0)   =   "0070001.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(16)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(15)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LabelInfo(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "mebBoekDatum"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TekstInfo(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "PolisDetail"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Annuleren"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Schoonvegen"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Afsluiten"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "KeuzeInfo(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "KeuzeInfo(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cbBoekDatum"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "TB2"
      TabPicture(1)   =   "0070001.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cbHerboeking"
      Tab(1).Control(1)=   "cbPolisHistoriek"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "rtbIndentTB2"
      Tab(1).ControlCount=   3
      Begin VB.CheckBox cbHerboeking 
         Alignment       =   1  'Right Justify
         Caption         =   "Voor Herboeking"
         Height          =   375
         Left            =   -68280
         TabIndex        =   18
         Top             =   120
         Width           =   1815
      End
      Begin VB.CheckBox cbBoekDatum 
         Caption         =   "Met Boekdatum"
         Height          =   255
         Left            =   5400
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox KeuzeInfo 
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
         TabIndex        =   10
         Top             =   810
         Width           =   5175
      End
      Begin VB.ComboBox KeuzeInfo 
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
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   120
         Width           =   5145
      End
      Begin VB.CommandButton Afsluiten 
         Caption         =   "&Boeken"
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
         Height          =   330
         Left            =   7230
         TabIndex        =   8
         Top             =   120
         Width           =   1320
      End
      Begin VB.CommandButton Schoonvegen 
         Caption         =   "Schoo&n"
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
         Left            =   7230
         TabIndex        =   7
         Top             =   480
         Width           =   1320
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
         Height          =   330
         Left            =   7230
         TabIndex        =   6
         Top             =   840
         Width           =   1320
      End
      Begin VB.ComboBox cbPolisHistoriek 
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
         Left            =   -74880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   120
         Width           =   4815
      End
      Begin RichTextLib.RichTextBox rtbIndentTB2 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   1
         Top             =   480
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   7646
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"0070001.frx":0038
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
      Begin MSFlexGridLib.MSFlexGrid PolisDetail 
         Height          =   3435
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   6059
         _Version        =   393216
         FixedCols       =   0
         FocusRect       =   2
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
      Begin MSMask.MaskEdBox TekstInfo 
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   11
         Top             =   450
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   16777215
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mebBoekDatum 
         Height          =   255
         Left            =   5760
         TabIndex        =   17
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   7
         Format          =   "MM/YYYY"
         Mask            =   "##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label LabelInfo 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Index           =   2
         Left            =   3660
         TabIndex        =   14
         Top             =   480
         Width           =   1560
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Boek&datum"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Boek.Rek."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   16
         Left            =   2640
         TabIndex        =   12
         Top             =   480
         Width           =   990
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Koers"
         Height          =   240
         Index           =   9
         Left            =   -71040
         TabIndex        =   4
         Top             =   4080
         Width           =   525
      End
      Begin VB.Label Label1 
         Caption         =   "Vervaldag"
         Height          =   240
         Index           =   6
         Left            =   -74910
         TabIndex        =   3
         Top             =   3330
         Width           =   1035
      End
   End
End
Attribute VB_Name = "KwijtingBoeken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim dokumentSleutel As String * 11
Dim KlantenRekening As String * 7

Dim RekBeheersKost As String * 7

Dim rsTB2 As ADODB.Recordset

Sub VraagWijzigingBoeken()
Dim TempoString As String
Dim TempoString9 As String

Dim Maatschappij As String * 4
Dim Operatie As Integer
Dim KopijString As String

MsgBox GridTextPolis
Msg = "Wijzigingen in rekeninguittreksels invoegen" + vbCrLf
Msg = Msg + "Bent U zeker ?"
Ktrl = MsgBox(Msg, vbYesNo + vbDefaultButton1 + vbQuestion)
If Ktrl = vbNo Then
    Exit Sub
Else
    TempoString = Mid(GridTextPolis, 1, InStr(GridTextPolis, vbCrLf) - 1)
    Maatschappij = Mid(TempoString, 1, InStr(TempoString, vbTab) - 1)
    KeuzeInfo(0).ListIndex = -1
    For Tel = 0 To KeuzeInfo(0).ListCount - 1
        If Maatschappij = Left(KeuzeInfo(0).List(Tel), 4) Then
            KeuzeInfo(0).Enabled = False
            KeuzeInfo(0).ListIndex = Tel
            Exit For
        End If
    Next
    If KeuzeInfo(0).ListIndex = -1 Then MsgBox "Maatschappij " + Maatschappij + " niet aanwezig !:exit sub"
    
    Operatie = Val(Right(TempoString, 1))
    KeuzeInfo(1).ListIndex = Operatie
    KeuzeInfo(1).Enabled = False
    'Eerste erbijvoegen
    GoSub VoegHetErBij
       
    Do
        GridTextPolis = Right(GridTextPolis, Len(GridTextPolis) - (Len(TempoString) + 2))
        If GridTextPolis = "" Then
            GridTextPolis = KopijString
            Exit Do
        Else
            TempoString = Mid(GridTextPolis, 1, InStr(GridTextPolis, vbCrLf) - 1)
            If Left(TempoString, 4) = Maatschappij And Operatie = Val(Right(TempoString, 1)) Then
                GoSub VoegHetErBij
            Else
                KopijString = KopijString + TempoString + vbCrLf
            End If
        End If
        
    Loop
    PolisDetail.Enabled = True
    TekstInfo(0).text = Left(BYPERDAT.PeriodeBoekjaar.text, 10)
    'TekstInfo(0).Enabled = False
    'KeuzeInfo(0).ListIndex = 0
    Afsluiten.Enabled = True
    Schoonvegen.Enabled = False
End If
Exit Sub

VoegHetErBij:
TempoString9 = Mid(TempoString, InStr(TempoString, vbTab) + 1)
TempoString9 = Left(TempoString9, Len(TempoString9) - 2)
PolisDetail.AddItem TempoString9, PolisDetail.Rows - 1
Return
    
End Sub

Private Sub Afsluiten_Click()
Dim Msg As String
Dim aa As String * 30
Dim TotaalPremie As Currency
Dim TotaalLoon As Currency
Dim T As Integer
Dim DummySleutel As String

If dokumentSleutel <> SleutelDok(121) Then
    Msg = dokumentSleutel + " <> " + SleutelDok(121) + vbCrLf + vbCrLf
    Msg = Msg + "MOGELIJKHEID 1: Teller is identiek, boekjaar is hoger/lager." + vbCrLf
    Msg = Msg + "U hebt dus het actief boekjaar of periode gewijzigd tijdens de aanmaak van dit dokument.  Probeer nogmaals NA KONTROLE." + vbCrLf + vbCrLf
    Msg = Msg + "MOGELIJKHEID 2: Boekjaar is identiek, teller is hoger/lager." + vbCrLf
    Msg = Msg + "Een andere gebruiker heeft ondertussen een dokument verwerkt." + vbCrLf + vbCrLf
    Msg = Msg + "Kontroleer eerst eens vooraleer de boeking nogmaals uit te voeren a.u.b. !!!"
    MsgBox Msg
    dokumentSleutel = SleutelDok(121)
    KwijtingBoeken.Caption = "Kwijtschriften beginnende (" + dokumentSleutel + ")"
    Exit Sub
End If

If Not DATE_CHECK((TekstInfo(0).text), PERIODAS_TEXT) Then
    Beep
    On Local Error Resume Next
    BYPERDAT.WindowState = 0
    BYPERDAT.PeriodeBoekjaar.SetFocus
    Exit Sub
End If

For T = 1 To PolisDetail.Rows - 2
    PolisDetail.Row = T
    PolisDetail.Col = 2
    TotaalPremie = TotaalPremie + Val(PolisDetail.text)
    PolisDetail.Col = 3
    TotaalLoon = TotaalLoon + Val(PolisDetail.text)
    PolisDetail.Col = 4
    If InStr(PolisDetail.text, "Kontroleer") Then
        MsgBox "Lijn " + Format(PolisDetail.Row - 2) + " kontroleren a.u.b. !"
        Exit Sub
    End If
Next

      Msg = RTrim$(KeuzeInfo(1).text) + " van" + vbCrLf
Msg = Msg + RTrim$(KeuzeInfo(0).text) + vbCrLf + vbCrLf
Msg = Msg + "Totaal premie(s) : " + Dec$((TotaalPremie), MASK_EURBH) + vbCrLf
Msg = Msg + "Totaal loon : " + Dec$((TotaalLoon), MASK_EURBH) + vbCrLf
Msg = Msg + "Boekdatum : " + TekstInfo(0).text
Ktrl = MsgBox(Msg, 292, "Afboeken " + Right(dokumentSleutel, 5) + " tot " + Format(Val(Right(dokumentSleutel, 5)) + PolisDetail.Rows - 3, "00000") + " ?")
If Ktrl = vbNo Then
    Exit Sub
End If

Ktrl = bOpen(TABLE_LEDGERACCOUNTS)
Ktrl = bOpen(TABLE_JOURNAL)
Ktrl = bOpen(TABLE_INVOICES)
bBegin
If WegBoekFout() Then
    bAbort
    Exit Sub
Else
    bEnd
    SS99 Right(dokumentSleutel, 5), 121
    If Dir$(LOCATION_ASWEB + "AS1.RCV") <> "" Then
        RasterSchoon
        Annuleren_Click
        Exit Sub
    Else
        dokumentSleutel = SleutelDok(121)
        KwijtingBoeken.Caption = "Kwijtschriften beginnende (" + dokumentSleutel + ")"
        SchoonVegen_Click
    End If
End If
If GridTextPolis <> "" Then
    VraagWijzigingBoeken
End If

End Sub

Private Sub Annuleren_Click()

If PolisDetail.Rows > 2 Then
    Msg$ = "Aangeduide verrichtingen negeren." + vbCrLf + vbCrLf + "Bent U zeker ?"
    Ktrl = MsgBox(Msg$, 292, "Voortijdig stoppen...")
    If Ktrl = 6 Then
    Else
        Exit Sub
    End If
End If
GridText = ""
Unload Me

End Sub

Private Sub cbPolisHistoriek_Click()

    Dim MAPIString As String
        
    Screen.MousePointer = vbHourglass
    rsTB2.AbsolutePosition = Me.cbPolisHistoriek.ListIndex + 1
    MAPIString = rsTB2("memoTB2").Value
    If Me.cbHerboeking.Value = 1 Then
        strTELEBIBIO = MAPIString
        MsgBox strTELEBIBIO
    Else
        strTELEBIBIO = ""
    End If
    Me.rtbIndentTB2.text = tb2Indent(MAPIString)
    Screen.MousePointer = vbNormal
    
End Sub


Private Sub cbTB2ktrl_Click()

    If Me.cbTB2ktrl.Value = vbChecked Then
        Me.cbBoekDatum.Visible = True
        Me.mebBoekDatum.Visible = True
        Me.mebBoekDatum.Enabled = True
        If Me.mebBoekDatum.ClipText = "" Then
            Me.mebBoekDatum.text = Format(Now, "mm/yyyy")
            Me.cbBoekDatum.Value = vbChecked
        End If
    Else
        Me.cbBoekDatum.Visible = False
        Me.mebBoekDatum.Visible = False
    End If

Me.cbPolisHistoriek.Clear
SSTab2.TabCaption(1) = "&TB2= (0)"
Me.rtbIndentTB2.text = ""

Dim MsgNoZero As String


Select Case cbTB2ktrl.Value
    Case vbUnchecked
        'niks
        
    Case vbChecked
       'nu eerst nog TB2controle
       Screen.MousePointer = vbHourglass
        If rsTB2.State = adStateOpen Then
            rsTB2.Close
        End If
        Msg = "SELECT * FROM TB2 WHERE memoTB2 Like '%PTY+006+" & Left(KeuzeInfo(0).text, 4) & "+%'"
        MsgNoZero = "SELECT * FROM TB2 WHERE memoTB2 Like '%PTY+006+" & Trim(Str(Val(Left(KeuzeInfo(0).text, 4)))) & "+%'"
        'PTY+006+0739+
        If Me.cbBoekDatum.Value = vbChecked Then
            Msg = Msg & " AND memoTB2 Like '%DTM+005:" & Me.mebBoekDatum.ClipText & ":005" & "%'"
            MsgNoZero = MsgNoZero & " AND memoTB2 Like '%DTM+005:" & Me.mebBoekDatum.ClipText & ":005" & "%'"
        End If
       
        Me.cbPolisHistoriek.Clear
        rsTB2.Open Msg, adntDB, adOpenDynamic, adLockOptimistic, adCmdText
        If rsTB2.RecordCount > 0 Then
            SSTab2.TabCaption(1) = "&TB2= (" & Trim(Str(rsTB2.RecordCount)) & ")"
            Do While Not rsTB2.EOF
                'strTB2 = rsTB2("memoTB2").Value
                Me.cbPolisHistoriek.AddItem "DocType=" & rsTB2("DocType").Value & " Aanvang=" & rsTB2("DatumAanvang").Value
                rsTB2.MoveNext
            Loop
        Else
            If rsTB2.State = adStateOpen Then
                rsTB2.Close
            End If
            rsTB2.Open MsgNoZero, adntDB, adOpenDynamic, adLockOptimistic, adCmdText
            If rsTB2.RecordCount > 0 Then
                SSTab2.TabCaption(1) = "&TB2= (" & Trim(Str(rsTB2.RecordCount)) & ")"
                Do While Not rsTB2.EOF
                    'strTB2 = rsTB2("memoTB2").Value
                    Me.cbPolisHistoriek.AddItem "DocType=" & rsTB2("DocType").Value & " Aanvang=" & rsTB2("DatumAanvang").Value
                    rsTB2.MoveNext
                Loop
            End If
        End If
        Screen.MousePointer = vbNormal
End Select
        
End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

  Set rsTB2 = New ADODB.Recordset
    rsTB2.CursorLocation = adUseClient
    
    
Ktrl = ktrlRVmemo
    
    
Dim LijnPos As Integer
Dim TotaalPremie As Currency
Dim TotaalLoon As Currency

Dim LijnErBij As String
Dim BeheerForfait As Currency
Dim BedragFee As Currency
Dim tbBeheerForfait As Currency

PolisDetail.Cols = 7
With PolisDetail
    .Col = 0
    .Row = 0
    .text = "Polisnummer"
    .ColWidth(0) = 1215
    
    .Col = 1
    .text = "Datum"
    .ColWidth(1) = 1005
    
    .Col = 2
    .text = "Totale premie"
    .ColWidth(2) = 1035
    .ColAlignment(2) = 1
        
    .Col = 3
    .text = "Loon"
    .ColWidth(3) = 795
    .ColAlignment(3) = 1
    
    .Col = 4
    .text = "Klant (Naam 1)"
    .ColWidth(4) = 1600
    
    .Col = 5
    .text = "XML/TLB2 Snippet"
    .ColWidth(5) = 1095
    
    .Col = 6
    .text = "Fee"
    .ColWidth(6) = 1035
    
End With

XDoEvents = DoEvents
Screen.MousePointer = vbHourglass
InstallMij

On Error Resume Next
RekBeheersKost = String99(READING, 289)
On Error Resume Next
tbBeheerForfait = String99(READING, 288)

KeuzeInfo(0).ListIndex = 0
KeuzeInfo(1).AddItem "0: Kies eerst bewerkingskode !"
KeuzeInfo(1).AddItem "1: Termijn"
KeuzeInfo(1).AddItem "2: Kontant"
KeuzeInfo(1).AddItem "3: Terugbetaling van de premies"
KeuzeInfo(1).AddItem "4: Vernietiging van verrichting"
KeuzeInfo(1).AddItem "5: Terugzending van kwitanties"
KeuzeInfo(1).AddItem "6: Commissielonen"
KeuzeInfo(1).AddItem "7: Schadegevallen"
KeuzeInfo(0).ListIndex = 0
KeuzeInfo(1).ListIndex = 0
dokumentSleutel = SleutelDok(121)
KwijtingBoeken.Caption = "Kwijtschriften beginnende (" + dokumentSleutel + ")"

KlantenRekening = LTrim$(RTrim$(String99(READING, 9)))

Dim arrayGTI() As String
Dim arrayGridLines() As String
Dim arrayGridCols() As String
Dim COUNT_TOLines As Integer
Dim COUNT_TOCols As Integer

If GridText <> "" Then
    arrayGTI = Split(GridTextIs, vbTab)
    Select Case arrayGTI(0)
        Case "001"
            KeuzeInfo(1).ListIndex = 1
        Case "002"
            KeuzeInfo(1).ListIndex = 2
        Case "003"
            KeuzeInfo(1).ListIndex = 3
        Case "006"
            KeuzeInfo(1).ListIndex = 6
        Case Else
            MsgBox "Typeboeking logica nog te voorzien voor " & arrayGTI(0)
    End Select
    arrayGridLines = Split(GridText, Chr$(13))
    For COUNT_TOLines = 0 To UBound(arrayGridLines) - 1
        arrayGridCols = Split(arrayGridLines(COUNT_TOLines), vbTab)
        TotaalPremie = TotaalPremie + Val(arrayGridCols(2))
        TotaalLoon = TotaalLoon + Val(arrayGridCols(3))
        If KeuzeInfo(1).ListIndex = 1 Then
            BeheerForfait = tbBeheerForfait
        Else
            beheersforfait = 0
        End If
        bGet TABLE_CONTRACTS, 0, Trim(arrayGridCols(0))
            If Ktrl Then
                MsgBox "logicafout", vbCritical
            Else
                RecordToVeld TABLE_CONTRACTS
            End If
            
            'Ktrl = BeheersKost(Trim(Mid(GridText, 1, 12)), BeheerForfait)
            LijnErBij = arrayGridCols(0) & vbTab & arrayGridCols(1) & vbTab & arrayGridCols(2) & vbTab & arrayGridCols(3) & vbTab & arrayGridCols(4) & vbTab & arrayGridCols(5)
            
            If KeuzeInfo(1).ListIndex = 6 Then
                LijnErBij = LijnErBij & vbTab & "0"
            Else
                If rsMAR(TABLE_CONTRACTS)("e070") = "1" Then
                    BedragFee = Val(arrayGridCols(6)) * Val(rsMAR(TABLE_CONTRACTS)("e071")) / 100
                    LijnErBij = LijnErBij & vbTab & Str(BedragFee)
                    'dit is een met FEE berekening!
                Else
                    If Trim(rsMAR(TABLE_CONTRACTS)("e069")) = "" Then
                        LijnErBij = LijnErBij & vbTab & BeheerForfait
                    Else
                        LijnErBij = LijnErBij & vbTab & Val(rsMAR(TABLE_CONTRACTS)("e069"))
                    End If
                End If
            End If
            
            PolisDetail.AddItem LijnErBij, PolisDetail.Rows - 1
            'GridText = Right(GridText, Len(GridText) - LijnPos)
    Next
        'LabelInfo(0).Caption = Dec$((TotaalPremie), "#########0")
    'LabelInfo(1).Caption = Dec$((TotaalLoon), "#########0")
    PolisDetail.Enabled = True
    If arrayGTI(2) = "Termijn" Then
        TekstInfo(0).text = Left(BYPERDAT.PeriodeBoekjaar.text, 10)
        TekstInfo(0).Enabled = False
    ElseIf arrayGTI(2) = "Commissies" Then
        TekstInfo(0).text = Right(BYPERDAT.PeriodeBoekjaar.text, 10)
        TekstInfo(0).Enabled = False
    Else
        TekstInfo(0).text = Left(BYPERDAT.PeriodeBoekjaar.text, 10)
        TekstInfo(0).Enabled = True
    End If
    
    KeuzeInfo(1).Enabled = False
    KeuzeInfo(0).ListIndex = 0
    Do While Left(KeuzeInfo(0).text, 4) <> "9000"
        If Mid(GridTextIs, 5, 4) = Left(KeuzeInfo(0).text, 4) Then
            KeuzeInfo(0).Enabled = False
            Exit Do
        Else
            KeuzeInfo(0).ListIndex = KeuzeInfo(0).ListIndex + 1
        End If
    Loop
    Afsluiten.Enabled = True
    Schoonvegen.Enabled = False
ElseIf GridTextPolis <> "" Then
    VraagWijzigingBoeken
Else
    Schoon
End If
Screen.MousePointer = vbNormal

End Sub

Private Sub Form_Unload(Cancel As Integer)

BYPERDAT.Boekjaar.Enabled = True
GridText = ""

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

End Sub

Private Sub InstallPolissen()
Dim Zoekstring As String * 24
Dim Maatschappij As String * 4
Dim Dummy As String * 30

Zoekstring = Mid(TekstInfo(0).text, 4, 2)
Maatschappij = Left(KeuzeInfo(0).text, 4)
bGetOrGreater TABLE_CONTRACTS, 3, Zoekstring
If Ktrl Or Left(KEY_BUF(TABLE_CONTRACTS), 2) <> Left(Zoekstring, 2) Then
    Beep
    Exit Sub
ElseIf Maatschappij = Mid(KEY_BUF(TABLE_CONTRACTS), 15, 4) Then
    RecordToVeld TABLE_CONTRACTS
    Select Case Left(vBibTekst(TABLE_CONTRACTS, "#vs97 #"), 1)
        Case "0", "7", "9"
        Case Else
            bGet TABLE_CUSTOMERS, 0, vBibTekst(TABLE_CONTRACTS, "#A110 #")
            If Ktrl Then
                Dummy = "KlantLink onmogelijk !!! Kontroleer !!!"
            Else
                RecordToVeld TABLE_CUSTOMERS
                Dummy = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
            End If
            PolisDetail.AddItem vBibTekst(TABLE_CONTRACTS, "#A000 #") & vbTab + Mid(vBibTekst(TABLE_CONTRACTS, "#AW_2 #"), 7, 2) + "/" + Mid(vBibTekst(TABLE_CONTRACTS, "#AW_2 #"), 5, 2) + "/" + Mid(PERIOD_FROMTO, 1, 4) & vbTab & vbTab & vbTab + Dummy & vbTab & 0, PolisDetail.Rows - 1
    End Select
End If
Screen.MousePointer = vbHourglass

Do
    bNext TABLE_CONTRACTS
    If Ktrl Or Left(KEY_BUF(TABLE_CONTRACTS), 2) <> Left(Zoekstring, 2) Then
        Screen.MousePointer = vbNormal
        Exit Do
    ElseIf Maatschappij = Mid(KEY_BUF(TABLE_CONTRACTS), 15, 4) Then
        RecordToVeld TABLE_CONTRACTS
        Select Case Left(vBibTekst(TABLE_CONTRACTS, "#vs97 #"), 1)
            Case "0", "7", "9"
            Case Else
            bGet TABLE_CUSTOMERS, 0, vBibTekst(TABLE_CONTRACTS, "#A110 #")
            If Ktrl Then
                Dummy = "KlantLink onmogelijk !!! Kontroleer !!!"
            Else
                RecordToVeld TABLE_CUSTOMERS
                Dummy = vBibTekst(TABLE_CUSTOMERS, "#A100 #")
            End If
            PolisDetail.AddItem vBibTekst(TABLE_CONTRACTS, "#A000 #") & vbTab + Mid(vBibTekst(TABLE_CONTRACTS, "#AW_2 #"), 7, 2) + "/" + Mid(vBibTekst(TABLE_CONTRACTS, "#AW_2 #"), 5, 2) + "/" + Mid(PERIOD_FROMTO, 1, 4) & vbTab & vbTab & vbTab + Dummy & vbTab & 0, PolisDetail.Rows - 1
        End Select
    End If
Loop

End Sub

Private Sub KeuzeInfo_Click(Index As Integer)
Dim T As Integer

Select Case Index
    Case 0
        bGet TABLE_SUPPLIERS, 0, vSet("CO" + Left(KeuzeInfo(0).text, 4), 12)
        If Ktrl Then
            'MsgBox "stop"
        Else
            RecordToVeld TABLE_SUPPLIERS
        End If
        bGet TABLE_LEDGERACCOUNTS, 0, vSet(vBibTekst(TABLE_SUPPLIERS, "#v016 #"), 7)
        If Ktrl Then
            MsgBox "stop voor rekening aan te maken!"
            TLB_RECORD(TABLE_LEDGERACCOUNTS) = ""
            vBib TABLE_LEDGERACCOUNTS, "700" + Left(KeuzeInfo(0).text, 4), "v019"
            vBib TABLE_LEDGERACCOUNTS, vBibTekst(TABLE_SUPPLIERS, "#A100 #"), "v020"
            bInsert TABLE_LEDGERACCOUNTS, 0
            TLB_RECORD(TABLE_LEDGERACCOUNTS) = ""
            vBib TABLE_LEDGERACCOUNTS, "440" + Left(KeuzeInfo(0).text, 4), "v019"
            vBib TABLE_LEDGERACCOUNTS, vBibTekst(TABLE_SUPPLIERS, "#A100 #"), "v020"
            'FVT(TABLE_LEDGERACCOUNTS, 0) = "440" + Left(KeuzeInfo(0).Text, 4)
            'FVT(TABLE_LEDGERACCOUNTS, 1) = FVT(TABLE_SUPPLIERS, 1)
            bInsert TABLE_LEDGERACCOUNTS, 0
            vBib TABLE_SUPPLIERS, vBibTekst(TABLE_LEDGERACCOUNTS, "#v019 #"), "v016"
            'FVT(TABLE_SUPPLIERS, 16) = FVT(TABLE_LEDGERACCOUNTS, 0)
            bUpdate TABLE_SUPPLIERS, 0
            LabelInfo(2).Caption = vBibTekst(TABLE_SUPPLIERS, "#v016 #")
        Else
            LabelInfo(2).Caption = vBibTekst(TABLE_SUPPLIERS, "#v016 #")
        End If
        If Me.cbTB2ktrl.Value = vbChecked Then
            cbTB2ktrl_Click
        End If
        
    Case 1
        If KeuzeInfo(1).ListIndex = 0 Then
            PolisDetail.Enabled = False
        Else
            PolisDetail.Enabled = True
        End If
End Select

End Sub

Private Sub KeuzeInfo_LostFocus(Index As Integer)

Select Case Index
    Case 1
        If KeuzeInfo(1).ListIndex = 1 And PolisDetail.Rows = 2 Then
            Msg$ = "Bestaande contracten met vervaldag voor de" + vbCrLf + Mid(TekstInfo(0).text, 4, 2) + "e maand inladen ?"
            Ktrl = MsgBox(Msg$, 292, "Polissen Termijn inladen")
            If Ktrl = 6 Then
                InstallPolissen
            End If
        End If
End Select
End Sub

Private Sub PolisDetail_Click()

On Local Error Resume Next
With PolisDetail
    .ColSel = .Cols - 1
'    .SelEndCol = 4
'    .SelStartRow = .Row
'    .SelEndRow = .Row
End With

End Sub

Private Sub PolisDetail_GotFocus()

SnelHelpPrint "[+] lijn bijvoegen, [-] verwijderen, [Enter] om te wijzigen", BL_LOGGING

End Sub

Private Sub PolisDetail_KeyPress(KeyAscii As Integer)
Dim Positie As Integer
Dim TempoPremie As Long
Dim TempoLoon As Long

Positie = PolisDetail.Row
Select Case KeyAscii
    Case 13, 43
        If KeyAscii = 43 Then
            PolisDetail.AddItem vbTab & vbTab & vbTab & vbTab & vbTab, PolisDetail.Rows - 1
            PolisDetail.Row = PolisDetail.Rows - 2
        ElseIf KeyAscii = 13 Then
            If Positie = PolisDetail.Rows - 1 Then
                MsgBox "Druk [+] toets om bij te voegen !"
                Exit Sub
            End If
        End If
        Load BoekKwijtingEdit
        BoekKwijtingEdit.Caption = Left(KeuzeInfo(0).text, 4) + ":" + KwijtingEdit.Caption
        BoekKwijtingEdit.Show 1
        KwijtingBoeken.SetFocus
        If KeyAscii = 43 And GridText = "ESC" Then
            PolisDetail.RemoveItem PolisDetail.Rows - 1
        End If
                        
    Case 43
    
    Case 45, 127
        If PolisDetail.Rows = 2 Then
            Exit Sub
        ElseIf Positie < 0 Then
            MsgBox "Eerst een lijn selekteren !", 0, "Lijn wijzigen"
            Exit Sub
        End If
        PolisDetail.RemoveItem Positie
        If PolisDetail.Rows - 1 > Positie + 1 Then
            PolisDetail.Row = Positie - 1
        End If
        KwijtingBoeken!PolisDetail.SetFocus
    Case Else
End Select

If PolisDetail.Rows = 2 Then
    Afsluiten.Enabled = False
Else
    Afsluiten.Enabled = True
End If

End Sub

Private Sub RasterSchoon()

With PolisDetail
    .Rows = 2
    .Row = 1
    .Col = 0
    .text = ""
    
    .Col = 1
    .text = ""
    
    .Col = 2
    .text = ""
    
    .Col = 3
    .text = ""
    
    .Col = 4
    .text = ""

    .Col = 5
    .text = ""

    .Col = 6
    .text = ""

End With

End Sub

Private Sub Schoon()
Dim T As Integer

Afsluiten.Enabled = False
TekstInfo(0).text = MIM_GLOBAL_DATE
RasterSchoon

End Sub

Private Sub PolisDetail_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 37 To 40
        PolisDetail_Click
End Select

End Sub


Private Sub SchoonVegen_Click()

Schoon

End Sub



Private Sub TekstInfo_GotFocus(Index As Integer)

TekstInfo(Index).SelStart = 0
TekstInfo(Index).SelLength = Len(TekstInfo(Index).text)

End Sub

Private Sub TekstInfo_LostFocus(Index As Integer)

Select Case Index
    Case 0
        If DATE_INVALID((TekstInfo(0).text)) Then
            TekstInfo(0).text = MIM_GLOBAL_DATE
            Beep
            TekstInfo(0).SetFocus
        ElseIf Not DATE_CHECK((TekstInfo(0).text), PERIODAS_TEXT) Then
            Beep
            BYPERDAT.WindowState = 0
            'BYPERDAT.PeriodeBoekjaar.SetFocus
        End If
End Select

End Sub

Private Function WegBoekFout%()
Dim BakdokumentSleutel As String * 11
Dim TempoDokSleutel As String
Dim dTotaalMij As Double
Dim dTotaalLoon As Double
Dim dTotaalBKost As Currency
Dim BedragTotaal As Currency

Dim BedragBeheer As Currency
Dim TotaalBedragBeheer As Currency

Dim Temporek As String * 7

'On Local Error GoTo IetsGingFout

WegBoekFout = False
DKTRL_CUMUL = 0: DKTRL_BEF = 0: DKTRL_EUR = 0
BakdokumentSleutel = dokumentSleutel
Unload frmBoeking
'frmBoeking.Hide
TLB_RECORD(TABLE_INVOICES) = ""
TLB_RECORD(TABLE_JOURNAL) = ""
'bijgevoegd voor scanning
If rsMAR(TABLE_INVOICES).State = adStateClosed Then
    Ktrl = bOpen(TABLE_INVOICES)
End If

dTotaalMij = 0
dTotaalLoon = 0

Screen.MousePointer = vbHourglass
For T = 1 To PolisDetail.Rows - 2
    rsMAR(TABLE_INVOICES).AddNew
    If bhEuro Then
        vBib TABLE_INVOICES, "EUR", "vs03"
    Else
        vBib TABLE_INVOICES, "BEF", "vs03"
    End If
    vBib TABLE_INVOICES, "1", "v040"
    vBib TABLE_INVOICES, Left(KeuzeInfo(1).text, 1), "v041"
    
    PolisDetail.Row = T
    PolisDetail.Col = 0
    bGet TABLE_CONTRACTS, 0, vSet(PolisDetail.text, 12)
    If Ktrl Then
        MsgBox "Polisnummer " + PolisDetail.text + vbCrLf + vbCrLf + "niet te vinden !"
        WegBoekFout = True
        dokumentSleutel = BakdokumentSleutel
        Exit Function
    Else
        RecordToVeld TABLE_CONTRACTS
        bGet TABLE_CUSTOMERS, 0, vSet(vBibTekst(TABLE_CONTRACTS, "#A110 #"), 12)
        If Ktrl Then
            MsgBox "Klantnummer " + vBibTekst(TABLE_CONTRACTS, "#A110 #") + " in lijn " + Str$(PolisDetail.Row) + vbCrLf + vbCrLf + "niet te vinden !"
            WegBoekFout = True
            dokumentSleutel = BakdokumentSleutel
            Exit Function
        Else
            vBib TABLE_JOURNAL, "K" + vBibTekst(TABLE_CONTRACTS, "#A110 #"), "v034"
            vBib TABLE_INVOICES, "K" + vBibTekst(TABLE_CONTRACTS, "#A110 #"), "v034"
            vBib TABLE_INVOICES, (PolisDetail.text), "A000"
            vBib TABLE_JOURNAL, KlantenRekening, "v019"
            vBib TABLE_JOURNAL, DATE_KEY((TekstInfo(0).text)), "v066"
            vBib TABLE_JOURNAL, dokumentSleutel, "v033"
            vBib TABLE_INVOICES, dokumentSleutel, "v033"
            PolisDetail.Col = 1

            vBib TABLE_JOURNAL, DATE_KEY(PolisDetail.text), "v035"
            vBib TABLE_INVOICES, DATE_KEY(PolisDetail.text), "v035"
            vBib TABLE_INVOICES, DATE_KEY(PolisDetail.text), "v036"
            Select Case Left(KeuzeInfo(1).text, 1)
                Case "1", "2", "6"
                    If Left(KeuzeInfo(1).text, 1) = "1" Then
                        vBib TABLE_CONTRACTS, Mid(DATE_KEY(PolisDetail.text), 5, 2), "v164"
                        vBib TABLE_CONTRACTS, Mid(DATE_KEY(PolisDetail.text), 7, 2), "v165"
                    End If
                    PolisDetail.Col = 4
                    vBib TABLE_JOURNAL, PolisDetail.text, "v067"
                    
                    'nieuw voor beheer...
                    PolisDetail.Col = 6
                    If Trim$(PolisDetail.text) = "" Then
                        BedragBeheer = 0
                    Else
                        BedragBeheer = Round(Val(PolisDetail.text), 2)
                    End If
                    dTotaalBKost = dTotaalBKost + BedragBeheer
                    
                    vBib TABLE_INVOICES, Str(BedragBeheer), "e069"
                    vBib TABLE_CONTRACTS, Str(BedragBeheer), "e069"
                    
                    PolisDetail.Col = 5
                    'MsgBox "TODO: snipped document: " & PolisDetail.Text
                    If PolisDetail.text = "" Then
                        rsMAR(TABLE_INVOICES)("rvXMLTB2") = ""
                    Else
                        rsMAR(TABLE_INVOICES)("rvXMLTB2") = PolisDetail.text
                    End If
                    
                    PolisDetail.Col = 2
                    BedragTotaal = BedragBeheer + Val(PolisDetail.text)
                    vBib TABLE_JOURNAL, Str$(BedragTotaal), "v068"
                    vBib TABLE_CONTRACTS, PolisDetail.text, "B010"
                    vBib TABLE_INVOICES, PolisDetail.text, "B010"
                    vBib TABLE_INVOICES, Str$(BedragTotaal), "v249"
                                                            
                    vBib TABLE_JOURNAL, LabelInfo(2).Caption, "v069"
                    If BedragTotaal <> 0 Then
                        bInsert TABLE_JOURNAL, 0
                    End If
                                        
                    Temporek = LabelInfo(2).Caption
                    Mid(Temporek, 1, 2) = "70"
                    vBib TABLE_JOURNAL, Temporek, "v019"
                                        
                    PolisDetail.Col = 3
                    vBib TABLE_JOURNAL, Str$(-Val(PolisDetail.text)), "v068"
                    vBib TABLE_CONTRACTS, PolisDetail.text, "B014"
                    vBib TABLE_INVOICES, PolisDetail.text, "B014"
                    If Val(PolisDetail.text) <> 0 Then
                        If Left(KeuzeInfo(1).text, 1) = "1" Then
                            dTotaalLoon = dTotaalLoon - Val(PolisDetail.text)
                        Else
                            bInsert TABLE_JOURNAL, 0
                        End If
                    End If
                    vBib TABLE_JOURNAL, vBibTekst(TABLE_JOURNAL, "#v069 #"), "v019"
                    vBib TABLE_JOURNAL, " ", "v069"
                    vBib TABLE_JOURNAL, Left(KeuzeInfo(1).text, 1), "v041"
                    
                    vBib TABLE_JOURNAL, Str$(-(Val(vBibTekst(TABLE_INVOICES, "#B010 #")) - Val(vBibTekst(TABLE_INVOICES, "#B014 #")))), "v068"
                    If Val(vBibTekst(TABLE_JOURNAL, "#v068 #")) <> 0 Then
                        If Left(KeuzeInfo(1).text, 1) = "1" Then
                            dTotaalMij = dTotaalMij + Val(vBibTekst(TABLE_JOURNAL, "#v068 #"))
                        Else
                            bInsert TABLE_JOURNAL, 0
                        End If
                    End If
                    
                    If Left(KeuzeInfo(1).text, 1) = "1" Then
                        TotaalBedragBeheer = TotaalBedragBeheer + BedragBeheer
                        'MsgBox "stop"
                    Else
                        vBib TABLE_JOURNAL, vBibTekst(TABLE_JOURNAL, "#v019 #"), "v069"
                        vBib TABLE_JOURNAL, RekBeheersKost, "v019"
                        vBib TABLE_JOURNAL, Left(KeuzeInfo(1).text, 1), "v041"
                        vBib TABLE_JOURNAL, Str$(-BedragBeheer), "v068"
                        If Val(vBibTekst(TABLE_JOURNAL, "#v068 #")) <> 0 Then
                            bInsert TABLE_JOURNAL, 0
                        End If
                    End If
                    
                    vBib TABLE_JOURNAL, " ", "v041"
                    Select Case Left(vBibTekst(TABLE_CONTRACTS, "#vs97 #"), 1)
                        Case "2", "0"
                            
                        Case Else
                            vBib TABLE_CONTRACTS, "1", "v041"
                    End Select
                    
                    bUpdate TABLE_CONTRACTS, 0
                    bInsert TABLE_INVOICES, 0
                    If Ktrl Then
                        WegBoekFout = True
                        dokumentSleutel = BakdokumentSleutel
                        Exit Function
                    End If
                    Mid(dokumentSleutel, 7, 5) = Format(Val(Right(dokumentSleutel, 5)) + 1, "00000")

                Case "3", "4", "5"
                    PolisDetail.Col = 4
                    vBib TABLE_JOURNAL, PolisDetail.text, "v067"
                    
                    PolisDetail.Col = 6
                    If Trim$(PolisDetail.text) = "" Then
                        BedragBeheer = 0
                    Else
                        BedragBeheer = PolisDetail.text
                    End If
                    vBib TABLE_INVOICES, Str$(BedragBeheer), "e069"
                    vBib TABLE_CONTRACTS, Str$(BedragBeheer), "e069"
                    dTotaalBKost = dTotaalBKost - BedragBeheer
                                        
                    PolisDetail.Col = 5
                    'MsgBox "TODO: snipped document: " & PolisDetail.Text
                    If PolisDetail.text = "" Then
                        rsMAR(TABLE_INVOICES)("rvXMLTB2") = ""
                    Else
                        rsMAR(TABLE_INVOICES)("rvXMLTB2") = PolisDetail.text
                    End If
                                        
                    PolisDetail.Col = 2
                    BedragTotaal = BedragBeheer + Val(PolisDetail.text)
                    vBib TABLE_JOURNAL, Str$(-BedragTotaal), "v068"
                    vBib TABLE_INVOICES, Str$(-BedragTotaal), "v249"
                    vBib TABLE_CONTRACTS, PolisDetail.text, "B090"
                    vBib TABLE_INVOICES, PolisDetail.text, "B090"
                    
                    vBib TABLE_JOURNAL, LabelInfo(2).Caption, "v069"
                    If BedragTotaal <> 0 Then
                        bInsert TABLE_JOURNAL, 0
                    End If
                                        
                    Temporek = LabelInfo(2).Caption
                    Mid(Temporek, 1, 2) = "70"
                    vBib TABLE_JOURNAL, Temporek, "v019"
                    
                    PolisDetail.Col = 3
                    vBib TABLE_JOURNAL, PolisDetail.text, "v068"
                    vBib TABLE_CONTRACTS, PolisDetail.text, "B094"
                    vBib TABLE_INVOICES, PolisDetail.text, "B094"
                    If Val(PolisDetail.text) <> 0 Then
                        bInsert TABLE_JOURNAL, 0
                    End If
          
                    vBib TABLE_JOURNAL, vBibTekst(TABLE_JOURNAL, "#v069 #"), "v019"
                    vBib TABLE_JOURNAL, " ", "v069"
                    vBib TABLE_JOURNAL, Str$(Val(vBibTekst(TABLE_INVOICES, "#B090 #")) - Val(vBibTekst(TABLE_INVOICES, "#B094 #"))), "v068"
                    If Val(vBibTekst(TABLE_JOURNAL, "#v068 #")) <> 0 Then
                        bInsert TABLE_JOURNAL, 0
                    End If
                    
                    vBib TABLE_JOURNAL, vBibTekst(TABLE_JOURNAL, "#v019 #"), "v069"
                    vBib TABLE_JOURNAL, RekBeheersKost, "v019"
                    vBib TABLE_JOURNAL, Left(KeuzeInfo(1).text, 1), "v041"
                    vBib TABLE_JOURNAL, Str(BedragBeheer), "v068"
                    If Val(vBibTekst(TABLE_JOURNAL, "#v068 #")) <> 0 Then
                        bInsert TABLE_JOURNAL, 0
                    End If

                    
                    
                    
                    bInsert TABLE_INVOICES, 0
                    If Ktrl Then
                        WegBoekFout = True
                        dokumentSleutel = BakdokumentSleutel
                        Exit Function
                    End If
                    Mid(dokumentSleutel, 7, 5) = Format(Val(Right(dokumentSleutel, 5)) + 1, "00000")
                
                Case Else
                    MsgBox "Nog niet beschikbaar"
                    Beep
                    WegBoekFout = True
                    dokumentSleutel = BakdokumentSleutel
                    Exit Function
            End Select
        End If
    End If
NextLijn:
Next

If Left(KeuzeInfo(1).text, 1) = "1" Then
    vBib TABLE_JOURNAL, LabelInfo(2).Caption, "v019"
    TempoDokSleutel = dokumentSleutel
    Mid(TempoDokSleutel, 7, 5) = Format(Val(Right(dokumentSleutel, 5)) - 1, "00000")
    
    vBib TABLE_JOURNAL, TempoDokSleutel, "v033"
    vBib TABLE_JOURNAL, LabelInfo(2).Caption, "v069"

    Temporek = vBibTekst(TABLE_JOURNAL, "#v019 #")
    Mid(Temporek, 1, 2) = "70"
    vBib TABLE_JOURNAL, Temporek, "v019"
    vBib TABLE_JOURNAL, "Termijn", "v067"
    vBib TABLE_JOURNAL, Str$(dTotaalLoon), "v068"
    If Val(vBibTekst(TABLE_JOURNAL, "#v068 #")) <> 0 Then
        bInsert TABLE_JOURNAL, 0
    End If
    
    If TotaalBedragBeheer <> 0 Then
        vBib TABLE_JOURNAL, vBibTekst(TABLE_JOURNAL, "#v019 #"), "v069"
        vBib TABLE_JOURNAL, RekBeheersKost, "v019"
        vBib TABLE_JOURNAL, Left(KeuzeInfo(1).text, 1), "v041"
        vBib TABLE_JOURNAL, Str$(-TotaalBedragBeheer), "v068"
        If Val(vBibTekst(TABLE_JOURNAL, "#v068 #")) <> 0 Then
            bInsert TABLE_JOURNAL, 0
        End If
        TotaalBedragBeheer = 0
    End If

    vBib TABLE_JOURNAL, vBibTekst(TABLE_JOURNAL, "#v019 #"), "v069"
    vBib TABLE_JOURNAL, LabelInfo(2).Caption, "v019"
    vBib TABLE_JOURNAL, Left(KeuzeInfo(1).text, 1), "v041"
    vBib TABLE_JOURNAL, Str$(dTotaalMij), "v068"
    If Val(vBibTekst(TABLE_JOURNAL, "#v068 #")) <> 0 Then
        bInsert TABLE_JOURNAL, 0
    End If
End If

Screen.MousePointer = vbNormal
If DKTRL_CUMUL <> 0 Then
    frmBoeking.cmdBoeken.Enabled = False
    MsgBox "LogikaFout bij vierkantskontrole journaal." + vbCrLf + vbCrLf + "Deze verrichting wordt geannuleerd.  Controleer zelf eerst en/of raadpleeg ons."
    frmBoeking.Show 1
    GoTo IetsGingFout
Else
    Select Case Left(Mim.cmdWegBoekModus.text, 1)
        Case "0"
        Case "1"
            If DKTRL_BEF Or DKTRL_EUR Then
                frmBoeking.Show 1
            End If
        Case "2"
            frmBoeking.Show 1
        Case Else
            MsgBox "situatie..."
    End Select
    If DKTRL_CUMUL Then GoTo IetsGingFout
End If
Mid(dokumentSleutel, 7, 5) = Format(Val(Right(dokumentSleutel, 5)) - 1, "00000")
WegBoekFout = False
Exit Function

IetsGingFout:
    Screen.MousePointer = vbNormal
    WegBoekFout = True
    dokumentSleutel = BakdokumentSleutel
    Exit Function

End Function

Function BeheersKost(PolisNummer As String, Bedrag As Currency) As Boolean

    Dim BetalingStr As String
       
    BeheersKost = True
    bGet TABLE_INVOICES, 2, PolisNummer
    On Error Resume Next
    BetalingStr = Trim$(rsMAR(TABLE_INVOICES)("v038").Value)
    If Left(rsMAR(TABLE_CONTRACTS)("vs97"), 1) = "2" Then
        BeheersKost = False
    ElseIf Left(BetalingStr, 2) = "KA" Then
        BeheersKost = False
    End If
    
End Function

Function ktrlRVmemo() As Boolean

Dim ktrlString As String

Err = 0
On Error Resume Next

'nieuwe poging 08/2009

ktrlRVmemo = True
ktrlString = rsMAR(TABLE_INVOICES)("rvXMLTB2").Name

If Err = 3265 Then
    rsMAR(TABLE_INVOICES).Close
    Msg = "ALTER TABLE Dokumenten ADD COLUMN rvXMLTB2 MEMO;"
    Err = 0
    adntDB.Execute Msg
    If Err Then
        MsgBox "Foutmelding bron: " & Err.Source & vbCrLf & "Foutkodenummer: " & Err.Number & vbCrLf & vbCrLf & "Foutmelding omschrijving:" & vbCrLf & Err.Description
    Else
        MsgBox Msg, vbInformation, "Met succes"
    End If
    MsgBox "Belangrijk veld werd toegevoegd. Gelieve het bedrijf opnieuw te openen a.u.b."
    ktrlRVmemo = False
End If

End Function


