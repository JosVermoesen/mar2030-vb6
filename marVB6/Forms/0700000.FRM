VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form InbrengFinancieel 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ctrl+F3 Financiële verrichtingen"
   ClientHeight    =   7890
   ClientLeft      =   75
   ClientTop       =   375
   ClientWidth     =   8610
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
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7890
   ScaleWidth      =   8610
   Begin TabDlg.SSTab SSTab1 
      Height          =   7815
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   13785
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Manueel"
      TabPicture(0)   =   "0700000.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LabelInfo(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblInfo(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblInfo(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblInfo(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblInfo(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LabelInfo(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LabelInfo(11)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LabelInfo(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblInfo(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "LabelInfo(12)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LabelInfo(13)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblInfo(7)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "LabelInfo(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Datum"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "KeuzeInfo(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Volgende"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Afsluiten"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "FinancieelDetail"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Annuleren"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Struktuur"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TextBoxWarningTestCompany"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ButtonReadCamt053"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CheckBoxSepaViewer"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Full CODA (versie 2.1)"
      TabPicture(1)   =   "0700000.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "mfgLijst"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "ButtonTransfer"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "ButtonAssign"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton ButtonAssign 
         Caption         =   "Toewijzen"
         Height          =   495
         Left            =   -74880
         TabIndex        =   26
         Top             =   6840
         Width           =   1575
      End
      Begin VB.CommandButton ButtonTransfer 
         Caption         =   "Overnemen"
         Enabled         =   0   'False
         Height          =   495
         Left            =   -68160
         TabIndex        =   25
         Top             =   6840
         Width           =   1575
      End
      Begin VB.CheckBox CheckBoxSepaViewer 
         Caption         =   ".XDA Inkijken"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   24
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CommandButton ButtonReadCamt053 
         Caption         =   "ReadCamt053"
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
         Left            =   7080
         TabIndex        =   23
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TextBoxWarningTestCompany 
         BackColor       =   &H000000FF&
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   4200
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "Opgelet: Dit is een testbedrijf!"
         Top             =   7560
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.CommandButton Struktuur 
         Caption         =   "&Gestructureerde Verrichting"
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
         Left            =   4080
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   7080
         Width           =   2805
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
         Left            =   6960
         TabIndex        =   8
         Top             =   7080
         Width           =   1445
      End
      Begin VB.ListBox FinancieelDetail 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   5010
         Left            =   0
         TabIndex        =   4
         Top             =   1800
         Width           =   8475
      End
      Begin VB.CommandButton Afsluiten 
         Caption         =   "&Boeken"
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
         Left            =   5460
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1445
      End
      Begin VB.CommandButton Volgende 
         Caption         =   "Ma&nueel"
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
         Left            =   5460
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   720
         Width           =   1445
      End
      Begin VB.ComboBox KeuzeInfo 
         BackColor       =   &H00FFFFFF&
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
         Left            =   60
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   3855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mfgLijst 
         Height          =   6615
         Left            =   -74880
         TabIndex        =   6
         Top             =   120
         Width           =   8355
         _ExtentX        =   14737
         _ExtentY        =   11668
         _Version        =   393216
         BackColor       =   -2147483624
         ForeColor       =   0
         FixedCols       =   0
         FocusRect       =   2
         AllowUserResizing=   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   0
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSComCtl2.DTPicker Datum 
         Height          =   315
         Left            =   5460
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   16580611
         CurrentDate     =   36327
         MinDate         =   35796
      End
      Begin VB.Label LabelInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Document           TegenR.       Bedrag Omschrijving                  Fin.Kort."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   0
         TabIndex        =   21
         Top             =   1560
         Width           =   8475
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "EindSaldo BEF"
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
         Index           =   7
         Left            =   2760
         TabIndex        =   20
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label LabelInfo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   3960
         TabIndex        =   19
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label LabelInfo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   1380
         TabIndex        =   18
         Top             =   1080
         Width           =   1350
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BeginSaldo BEF"
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
         Index           =   6
         Left            =   60
         TabIndex        =   17
         Top             =   1080
         Width           =   1275
      End
      Begin VB.Label LabelInfo 
         Caption         =   "Datu&m uittreksel"
         Height          =   240
         Index           =   1
         Left            =   5460
         TabIndex        =   1
         Top             =   120
         Width           =   1440
      End
      Begin VB.Label LabelInfo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Index           =   11
         Left            =   3960
         TabIndex        =   16
         Top             =   420
         Width           =   1350
      End
      Begin VB.Label LabelInfo 
         Caption         =   "Financiële &Rekening"
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BeginSaldo EUR"
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
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   780
         Width           =   1275
      End
      Begin VB.Label lblInfo 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Eindsaldo EUR"
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
         Index           =   5
         Left            =   2760
         TabIndex        =   13
         Top             =   780
         Width           =   1185
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1380
         TabIndex        =   12
         Top             =   780
         Width           =   1350
      End
      Begin VB.Label lblInfo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3960
         TabIndex        =   11
         Top             =   780
         Width           =   1350
      End
      Begin VB.Label LabelInfo 
         Alignment       =   1  'Right Justify
         Caption         =   "Uittreksel"
         Height          =   240
         Index           =   3
         Left            =   4380
         TabIndex        =   10
         Top             =   120
         Width           =   900
      End
   End
End
Attribute VB_Name = "InbrengFinancieel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim rsAny As ADODB.Recordset

Dim BeginBalans As Integer
Dim dokumentSleutel As String * 8

Dim DefaultRekening As String * 7 'Leveranciers

Dim RekeningNummer(9) As String * 7
Dim Uittreksel(9) As String * 6
Dim RecNummer(9) As Integer

Dim BekomenKorting As String * 7
Dim ToegestaneKorting As String * 7


Dim bCtrl As Boolean
Dim COUNT_TO As Integer
Dim iOptelControle As Integer

'Begin
Dim sDatumAanmaak As String
Dim sToepassingsCode As String
Dim sNaamBestemmeling As String

'OudSaldo
Dim sRekeningNummer As String
Dim sUittreksel As String
Dim cOudSaldo As Currency
Dim sDatumOudSaldo As String
Dim sNaamRekeninghouder As String
Dim sOmschrijvingRekening As String

'BewegingsArtikel1
Dim sRefFinInstelling As String
Dim sRefFinInstelling2 As String
Dim sDC As String
Dim cBedrag As Currency
Dim sVerrichting As String
Dim sMededeling As String
Dim sMDDZone1 As String
Dim sMDDZone2 As String
Dim sBoekDatum As String
Dim sDagAfschriftVolgNummer As String
Dim sValutadatum As String

'BewegingsArtikel2
Dim sMededeling2 As String
Dim sRefKlant(1) As String
Dim sMuntVerrichting As String
Dim cBedragMunt As Currency

'BewegingsArtikel3
Dim sRekeningTP As String
Dim sITcodesTP As String
Dim sRekeningTPextra As String
Dim sNaamEnAdres(2) As String

'NieuwSaldo
Dim sUittreksel2 As String
Dim sRekeningNummer2 As String
Dim cNieuwSaldo As Currency
Dim sDatumNieuwSaldo As String

'EindOpname
Dim iOptelCtrlCheckUp As Integer
Dim cDebetSaldo As Currency
Dim cCreditSaldo As Currency

Private Function CtrlDocuments(isSeller As Boolean, refIsCor As Boolean, refString As String, IbanAccount As String, lineAmount As String) As String
    
    Dim VoorLetter  As String * 1
    Dim aa          As Variant
    Dim dBetaald    As Double
    Dim dTotaal     As Double
    Dim checkerSEPA As String
    Dim clientNumber As String
    Dim formattedAmount As String
    Dim documentFound As String
                        
    formattedAmount = Str(Val(lineAmount))
    If InStr(refString, "ORef: ") Then

        documentFound = Mid(refString, InStr(refString, "ORef: ") + 6)
    Else
        documentFound = ""
    End If
    
    CtrlDocuments = ""
    Set rsAny = New ADODB.Recordset
    
    ' v259 is banc account
    ' v249 te betalen
    ' refString starting with 104 is insurance
    ' clientNumber is 6 numbers on position 4 to 7 + position 9 to 10
    If isSeller Then
        If documentFound = "" Then
            Msg = "SELECT Leveranciers.A110, Leveranciers.A100, Leveranciers.v259, "
            Msg = Msg & "Dokumenten.v033, Dokumenten.v034, Dokumenten.v035, Dokumenten.v036, "
            Msg = Msg & "Dokumenten.v037, Dokumenten.v038, Dokumenten.v039, Dokumenten.v249, "
            Msg = Msg & "Dokumenten.v411, Dokumenten.rvDM, Dokumenten.rvID "
            Msg = Msg & "FROM Leveranciers, Dokumenten "
            Msg = Msg & "WHERE Dokumenten.v034 = 'L' + Leveranciers.A110 "
            Msg = Msg & "AND Dokumenten.v039 = '" + refString + "' "
            Msg = Msg & "AND Leveranciers.v259 = '" + IbanAccount + "' "
            Msg = Msg & "AND Str(Val(Dokumenten.v249)) = '" + formattedAmount + "' "
            Msg = Msg & "ORDER BY Dokumenten.v037 "
        Else
            Msg = "SELECT Leveranciers.A110, Leveranciers.A100, Leveranciers.v259, "
            Msg = Msg & "Dokumenten.v033, Dokumenten.v034, Dokumenten.v035, Dokumenten.v036, "
            Msg = Msg & "Dokumenten.v037, Dokumenten.v038, Dokumenten.v039, Dokumenten.v249, "
            Msg = Msg & "Dokumenten.v411, Dokumenten.rvDM, Dokumenten.rvID "
            Msg = Msg & "FROM Leveranciers, Dokumenten "
            Msg = Msg & "WHERE Dokumenten.v034 = 'L' + Leveranciers.A110 "
            Msg = Msg & "AND Dokumenten.v033 = '" + documentFound + "' "
        End If
    Else
        Msg = "SELECT Klanten.A110, Klanten.A100, Klanten.v259, "
        Msg = Msg & "Dokumenten.v033, Dokumenten.v034, Dokumenten.v035, Dokumenten.v036, "
        Msg = Msg & "Dokumenten.v037, Dokumenten.v038, Dokumenten.v039, Dokumenten.v249, "
        Msg = Msg & "Dokumenten.v411, Dokumenten.rvDM, Dokumenten.rvID "
        Msg = Msg & "FROM Klanten, Dokumenten "
        Msg = Msg & "WHERE Dokumenten.v034 = 'K' + Klanten.A110 "
        Msg = Msg & "AND Str(Val(Dokumenten.v249)) = '" + formattedAmount + "' "
        If refIsCor And Mid(refString, 1, 3) = "104" Then
            clientNumber = Mid(refString, 4, 4) + Mid(refString, 9, 2)
            Msg = Msg & "AND Klanten.A110 = '" + clientNumber + "' "
        End If
        Msg = Msg & "ORDER BY Dokumenten.v037 "
    End If
    Screen.MousePointer = vbHourglass
   
    On Error Resume Next
    Err = 0
    rsAny.CursorLocation = adUseClient
    
    SnelHelpPrint Msg, BL_LOGGING
    Screen.MousePointer = vbHourglass
    rsAny.Open Msg, adntDB, adOpenDynamic, adLockOptimistic
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    ElseIf rsAny.RecordCount = 0 Then
        MsgBox "Geen documenten gevonden"
        Exit Function
    ElseIf rsAny.RecordCount = 1 Then
        rsAny.MoveFirst
        CtrlDocuments = rsAny(3) & "|" & rsAny(1) & "|" & rsAny(0)
        GoTo FinishFunction
    Else
        rsAny.MoveFirst
        Stop
        Do While Not rsAny.EOF
            GoSub ValidateDocument
            rsAny.MoveNext
        Loop
        rsAny.Close
        Set rsAny = Nothing
    End If
    Screen.MousePointer = vbNormal
    
    Exit Function

ValidateDocument:
    Stop
    Return
    
FinishFunction:
    rsAny.Close
    Set rsAny = Nothing
    
End Function


Private Sub Afsluiten_Click()
Dim T               As Integer
Dim TempoTekst      As String
Dim DummySleutel    As String * 5

If Not DATE_CHECK(Format(Datum.Value, "dd/mm/yyyy"), PERIODAS_TEXT) Then
    Beep
    Datum.SetFocus
    Exit Sub
ElseIf FinancieelDetail.ListCount = 0 Then
    Beep
    MsgBox "Verrichtingen inbrengen a.u.b. !!!"
    Exit Sub
End If

bGet TABLE_LEDGERACCOUNTS, 0, Left(KeuzeInfo(0), 7)
If Ktrl Then
    MsgBox "onlogische situatie"
    Exit Sub
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
End If
bGet TABLE_JOURNAL, 2, UCase(Left(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), 2)) + Right(Format(Datum.Value, "dd/mm/yyyy"), 2) + Format(Val(LabelInfo(11).Caption) - 1, "0000")
If Ktrl Then
    MsgBox "Dit zou het eerste uittreksel binnen het WERKELIJK jaar moeten zijn...  Kontroleer eventueel"
Else
    RecordToVeld TABLE_JOURNAL
    If vBibTekst(TABLE_JOURNAL, "#v066 #") > Format(Datum.Value, "yyyymmdd") Then
        Msg = "Er zijn reeds uittreksels met een hogere datum !" + vbCrLf + vbCrLf
        Msg = Msg + "Laatste uittreksel nr. " + UCase$(Left(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), 2)) + Right(Format(Datum.Value, "dd/mm/yyyy"), 2) + Format(Val(LabelInfo(11).Caption) - 1, "0000") + " dateert van : " + DATE_TEXT(vBibTekst(TABLE_JOURNAL, "#v066 #")) + vbCrLf + vbCrLf
        Msg = Msg + "Vervolg.  Bent U zeker ?"
        Ktrl = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, "Uittreksel afsluiten")
        If Ktrl = vbNo Then
            Exit Sub
        End If
    End If
End If

If bhEuro Then
    Msg = "Datum uittreksel " & Datum.Value & " en bekomen eindsaldo EUR " & lblInfo(1).Caption & vbCrLf & vbCrLf & "Hierna wordt de boekhouding bijgewerkt.  Bent U zeker ?"
Else
    Msg = "Datum uittreksel " & Datum.Value & " en bekomen eindsaldo BEF " & LabelInfo(13).Caption & vbCrLf & vbCrLf & "Hierna wordt de boekhouding bijgewerkt.  Bent U zeker ?"
End If
Ktrl = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, "Uittreksel : " + UCase$(Left(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), 2)) + Right(Format(Datum.Value, "dd/mm/yyyy"), 2) + Format(Val(LabelInfo(11).Caption), "0000"))
If Ktrl = vbYes Then
    bBegin
    If WegBoekFout() Then
        bAbort
        Exit Sub
    Else
        bEnd
        DummySleutel = "s" + Format(RecNummer(KeuzeInfo(0).ListIndex), "000")
        bGet TABLE_COUNTERS, 0, DummySleutel
        If Ktrl Then
            MsgBox "TellerStop " & DummySleutel & ".  kontakteer R&&Vsoft"
        Else
            RecordToVeld TABLE_COUNTERS
            FL99_RECORD = Str$(Val(LabelInfo(11).Caption))
            If BAModus = 1 Then
                vBib TABLE_COUNTERS, FL99_RECORD, "v217 "
            Else
                vBib TABLE_COUNTERS, FL99_RECORD, DummySleutel
            End If
            bUpdate TABLE_COUNTERS, 0
            If Ktrl Then
                MsgBox "Update TellerStop " & DummySleutel & ".  kontakteer R&&Vsoft"
            End If
        End If
        
        DummySleutel = "s101"
        bGet TABLE_COUNTERS, 0, DummySleutel
        If Ktrl Then
            MsgBox "TellerStop.  Versiekonflikt !  Kontakteer R&&Vsoft"
        Else
            RecordToVeld TABLE_COUNTERS
            FL99_RECORD = Left(KeuzeInfo(0).text, 7)
            If BAModus = 1 Then
                vBib TABLE_COUNTERS, FL99_RECORD, "v217 "
            Else
                vBib TABLE_COUNTERS, FL99_RECORD, DummySleutel
            End If
            bUpdate TABLE_COUNTERS, 0
            If Ktrl Then
                MsgBox "UpdateStop Teller. kontakteer R&&Vsoft"
            End If
        End If
        bClose TABLE_COUNTERS
        Unload Me
        GridText = ""
    End If
End If

End Sub

Private Sub Annuleren_Click()

If FinancieelDetail.ListCount Then
    Msg$ = "Aangeduide verrichtingen negeren." + vbCrLf + vbCrLf + "Bent U zeker ?"
    Ktrl = MsgBox(Msg$, 292)
    If Ktrl = 6 Then
    Else
        Exit Sub
    End If
End If
Unload Me
GridText = ""

End Sub

Private Sub ButtonAssign_Click()

    MsgBox "TODO: Toewijzen"
    
End Sub

Private Sub ButtonReadCamt053_Click()

    'Exit Sub

    Dim filePath As String
    Dim actionResult As Boolean
    Dim qualityOfBeginning As Integer
    
    Dim xdaLocation As String
    
    xdaLocation = LCase(LaadTekst("dnnInstellingen", "CodaIOMap"))
    If xdaLocation = "" Then
        xdaLocation = LCase(LOCATION_DESKTOP)
    End If
        
    On Error GoTo OnCancel
    
    Mim.Teken.filename = ""
    Mim.Teken.InitDir = xdaLocation
    Mim.Teken.CancelError = True
    Mim.Teken.Filter = "Alle bestanden (*.xda)|*.xda"
    Mim.Teken.ShowOpen
    
    filePath = LCase(Mim.Teken.filename)
    
    Dim checkDirChange As String
    checkDirChange = LCase(Mid(Mim.Teken.filename, 1, InStr(Mim.Teken.filename, Mim.Teken.FileTitle) - 1))
    
    If xdaLocation <> checkDirChange Then
        If InStr(checkDirChange, "coda\in") <> 0 Then
            MsgBox "Inladen van verwerkte documenten is verboden", vbCritical
            Exit Sub
        End If
        
        Msg = ".XDA en .XML voor SEPA betaalbestanden locatie staat ingesteld op:" & vbCrLf
        Msg = Msg & xdaLocation & vbCrLf & vbCrLf
        Msg = Msg & "Mag de standaard locatie vanaf nu gewijzigd worden naar:" & vbCrLf
        Msg = Msg & checkDirChange
        Ktrl = MsgBox(Msg, vbYesNo + vbQuestion + vbDefaultButton2, "Uittreksel afsluiten")
        If Ktrl = vbYes Then
            BeWaarTekst "dnnInstellingen", "CodaIOMap", checkDirChange
            MsgBox "Herstart het inladen van het document", vbInformation
            Exit Sub
        Else
            MsgBox "Probeer opnieuw", vbExclamation
            Exit Sub
        End If
    End If
    
    If Me.CheckBoxSepaViewer.Value = vbChecked Then
        actionResult = ReadCamt053XDA(filePath, True)
    Else
        actionResult = ReadCamt053XDA(filePath, False)
    End If
    
    ReDim xdaOMSArray(50)
    ReDim xdaDATAArray(50)
    xdaOMSArray() = Split(xdaOMS, vbTab)
    xdaDATAArray() = Split(xdaDATA, vbTab)
            
    ReDim xdaLinesOMSArray(20)
    xdaLinesOMSArray() = Split(xdaLinesOMS, vbTab)
    xdaParseToArray (xdaLinesDATA)
    'For TTT = 0 To UBound(xdaLinesOMSArray): Print TTT; " "; xdaLinesOMSArray(TTT); " => "; xdaLinesDATAArray(0, TTT): Next
    
    Dim beginSaldoXDA As Double
    Dim bedragTransactie As Double
    Dim cumulSaldo As Double
           
    If xdaDATAArray(3) <> "" Then
        'check if number of account statement
        If Val(LabelInfo(11).Caption) = Val(xdaDATAArray(3)) Then 'ok
            qualityOfBeginning = qualityOfBeginning + 1
        Else
        
        End If
    End If
    
    beginSaldoXDA = Val(xdaDATAArray(6))
    If beginSaldoXDA <> Val(Dec((lblInfo(0)), "")) Then
    Else
        qualityOfBeginning = qualityOfBeginning + 1
    End If
    
    Dim aa As String
    Select Case qualityOfBeginning
        Case 0
            Msg = "Noch teller (niet essentiëel), noch beginsaldo (essentieel)" & vbCrLf
            Msg = Msg + "Bij beheer van meerdere rekeningen, duidt eerst de rekening aan voor .XDA import en probeer opnieuw"
            MsgBox Msg, vbInformation
            Exit Sub
        Case 1
            SnelHelpPrint "Beginsaldo ok", BL_LOGGING
        Case 2
            SnelHelpPrint "Volgnummer + beginsaldo ok", BL_LOGGING
    End Select
    
    aa = "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & "" & vbTab & beginSaldoXDA
    cumulSaldo = beginSaldoXDA
    mfgLijst.AddItem aa, mfgLijst.Rows - 1
        
    On Local Error GoTo 0
    SnelHelpPrint "Preparing book list with xdaLinesData", BL_LOGGING
    
    Dim resultReturn As String
    Dim isCor As Boolean
    Dim isSeller As Boolean
    Dim lineReference As String
    Dim resultArray() As String
    Dim amountToCheck As String
            
    For T = 0 To UBound(xdaLinesDATAArray) - 1
        aa = Dec(T + 1, "000") & vbTab
        If Trim(xdaLinesDATAArray(T, 4)) = "" Then
            bedragTransactie = Val(xdaLinesDATAArray(T, 1))
            amountToCheck = xdaLinesDATAArray(T, 1)
        Else
            bedragTransactie = Val(xdaLinesDATAArray(T, 4))
            amountToCheck = xdaLinesDATAArray(T, 4)
        End If
        
        'Msg = "Entry BBA code: " & xdaLinesDATAArray(T, 2) & vbCrLf
        'Msg = Msg & "Tx BBA code: " & xdaLinesDATAArray(T, 11)
        'MsgBox Msg, vbInformation
                
        aa = aa & xdaLinesDATAArray(T, 2) & vbTab
        'TODO: Evaluate Debit or Credit
        Select Case xdaLinesDATAArray(T, 2)
            Case "0101000" '(5-C) SEPA Credit Transfer Type SEPA Transfer with Structured Communiaction (Creditor) (SCOR)
                bedragTransactie = -bedragTransactie
                MsgBox "Stop"
                aa = aa & "-" & vbTab
                aa = aa & "??????" & vbTab
                aa = aa & Val(amountToCheck) & vbTab
                aa = aa & "" & vbTab
                
            Case "0103000" '(5-C) SEPA Credit Transfer Type SEPA Transfer (Creditor) (Ustrd)
                bedragTransactie = -bedragTransactie
                MsgBox "Stop"
                If Len(xdaLinesDATAArray(T, 9)) = 12 Then
                    isCor = True
                    lineReference = xdaLinesDATAArray(T, 9)
                Else
                    isCor = False
                    If Len(xdaLinesDATAArray(T, 10)) = 0 Then
                        MsgBox "Empty Ustrd, Logic?"
                        lineReference = "-"
                    Else
                        lineReference = xdaLinesDATAArray(T, 10)
                    End If
                End If
                                    
                ReDim resultArray(5)
                resultReturn = CtrlDocuments(True, isCor, lineReference, xdaLinesDATAArray(T, 7), amountToCheck)
                If resultReturn = "" Then
                    MsgBox "Geen resultaat", vbExclamation
                Else
                    resultArray() = Split(resultReturn, "|")
                    aa = aa & "+" & resultArray(0) & vbTab
                    aa = aa & String99(READING, 9) & vbTab
                    aa = aa & Val(amountToCheck) & vbTab
                    aa = aa & resultArray(1) & vbTab
                End If
                
                aa = aa & "-" & vbTab
                aa = aa & "??????" & vbTab
                aa = aa & Val(amountToCheck) & vbTab
                aa = aa & "" & vbTab
                                
            Case "0107000" '(5-C) SEPA Credit Transfer Type SEPA Transfer (Creditor) (Ustrd)
                                
                bedragTransactie = -bedragTransactie
                
                If Len(xdaLinesDATAArray(T, 9)) = 12 Then
                    isCor = True
                    lineReference = xdaLinesDATAArray(T, 9)
                Else
                    isCor = False
                    If Len(xdaLinesDATAArray(T, 10)) = 0 Then
                        MsgBox "Empty Ustrd, Logic?"
                        lineReference = "-"
                    Else
                        lineReference = xdaLinesDATAArray(T, 10)
                    End If
                End If
                                    
                ReDim resultArray(5)
                resultReturn = CtrlDocuments(True, isCor, lineReference, xdaLinesDATAArray(T, 7), amountToCheck)
                If resultReturn = "" Then
                    MsgBox "Geen resultaat", vbExclamation
                    aa = aa & "-" & vbTab
                    aa = aa & "??????" & vbTab
                    aa = aa & Val(amountToCheck) & vbTab
                    aa = aa & xdaLinesDATAArray(5, 1) & vbTab
                Else
                    resultArray() = Split(resultReturn, "|")
                    aa = aa & "-" & resultArray(0) & vbTab
                    aa = aa & String99(READING, 10) & vbTab
                    aa = aa & Val(amountToCheck) & vbTab
                    aa = aa & resultArray(1) & vbTab
                End If
                                                
            Case "0201000" '(5-C) Unknow BBA Type Incoming Transfer (Creditor)
                bedragTransactie = -bedragTransactie
                MsgBox "Stop"
                aa = aa & "-" & vbTab
                aa = aa & "??????" & vbTab
                aa = aa & Val(amountToCheck) & vbTab
                aa = aa & "" & vbTab
                                
            Case "0402000" '(5-C) ATM Cash Withdrawal Type Card payment
                bedragTransactie = -bedragTransactie
                MsgBox "Stop"
                aa = aa & "-" & vbTab
                aa = aa & "??????" & vbTab
                aa = aa & Val(amountToCheck) & vbTab
                aa = aa & xdaLinesDATAArray(T, 10) & vbTab
                                
            Case "0501000" '(5-C) SEPA Direct Debit CORE Type Direct Debit (Creditor) (Ustrd)
                bedragTransactie = -bedragTransactie
                MsgBox "Stop"
                aa = aa & "-" & vbTab
                aa = aa & "??????" & vbTab
                aa = aa & Val(amountToCheck) & vbTab
                aa = aa & xdaLinesDATAArray(T, 5) & vbTab
                                            
            Case "0503000"
                '(5-C) Unknown BBA Type Direct Debit (Debtor) (Ustrd) annulation?
                bedragTransactie = -bedragTransactie
                MsgBox "Stop"
                aa = aa & "-" & vbTab
                aa = aa & "??????" & vbTab
                aa = aa & Val(amountToCheck) & vbTab
                aa = aa & xdaLinesDATAArray(T, 6) & vbTab
                                            
            Case "8022000"
                '(5-C) Unknown BBA Type Direct Debit (Debtor) (Ustrd) annulation?
                bedragTransactie = -bedragTransactie
                MsgBox "Stop"
                aa = aa & "-" & vbTab
                aa = aa & "??????" & vbTab
                aa = aa & Val(amountToCheck) & vbTab
                aa = aa & xdaLinesDATAArray(T, 10) & vbTab
            
            Case "0150000", "0250000"
                '0150000 (5-D) Unknown BBA Type SEPA Transfer (Debtor) (SCOR/Ustrd)
                '0250000 (5-D) Unknown BBA Type Incoming Transfer (Debtor) (SCOR)
                If Len(xdaLinesDATAArray(T, 9)) = 12 Then
                    isCor = True
                    lineReference = xdaLinesDATAArray(T, 9)
                Else
                    isCor = False
                    If Len(xdaLinesDATAArray(T, 10)) = 0 Then
                        MsgBox "Empty Ustrd, Logic?"
                        lineReference = "-"
                    Else
                        lineReference = xdaLinesDATAArray(T, 10)
                    End If
                End If
                    
                If Len(lineReference) > 25 Then
                    aa = aa & "+" & vbTab
                    aa = aa & "??????" & vbTab
                    aa = aa & Val(amountToCheck) & vbTab
                    aa = aa & xdaLinesDATAArray(T, 6) & vbTab
                Else
                    ReDim resultArray(5)
                    resultReturn = CtrlDocuments(False, isCor, lineReference, xdaLinesDATAArray(T, 7), amountToCheck)
                    If resultReturn = "" Then
                        MsgBox "Geen resultaat", vbExclamation
                    Else
                        resultArray() = Split(resultReturn, "|")
                        aa = aa & "+" & resultArray(0) & vbTab
                        aa = aa & String99(READING, 9) & vbTab
                        aa = aa & Val(amountToCheck) & vbTab
                        aa = aa & resultArray(1) & vbTab
                    End If
                End If
            
            Case "0254000" '(5-D) Unknown BBA Code Incoming Transfer
                MsgBox "Stop"
                aa = aa & "+" & vbTab
                aa = aa & "??????" & vbTab
                aa = aa & Val(xdaLinesDATAArray(T, 1)) & vbTab
                aa = aa & "" & vbTab
                            
            Case "0550000"
                '(5-D) Unknown BBA Type Direct Debit (Debtor) (Ustrd)
                MsgBox "Stop"
                aa = aa & "+" & vbTab
                aa = aa & "??????" & vbTab
                aa = aa & Val(xdaLinesDATAArray(T, 1)) & vbTab
                aa = aa & "" & vbTab
                            
            Case Else
                MsgBox "BBA buiten controle: " & xdaLinesDATAArray(T, 2)
        End Select
        cumulSaldo = cumulSaldo + bedragTransactie
        aa = aa & cumulSaldo
        mfgLijst.AddItem aa, mfgLijst.Rows - 1
    Next
    
    MsgBox "TODO: give chance to edit for lines not 100% sure"
        
    MsgBox "TODO: book or [Esc]!"
    GoTo OnCancel
    If Not CopyFile(xdaLocation, LCase(LOCATION_COMPANYDATA + "coda\in"), Mim.Teken.FileTitle) Then
        MsgBox filePath & " kan niet verplaatst worden.  Probeer eventueel manueel", vbInformation
    Else
        Kill filePath
    End If
       
OnCancel:
    
End Sub

Function fnTestBetaalopdracht()

    MsgBox "Stop in ontwikkeling"
    Dim boString As String * 128
    Dim eoString As String * 128
    Dim gostring As String * 128
    
    Mid(boString, 1, 1) = "0"                       'beginopname
    Mid(boString, 4, 2) = "00"                      'voorwerp van betaling
    Mid(boString, 6, 6) = Format(Now, "ddmmyy")     'datum opmaak
    Mid(boString, 12, 3) = "123"                    'codenummer fin.instelling
    
End Function






Private Sub ButtonTransfer_Click()

    MsgBox "TODO: Overnemen"
    
End Sub

Private Sub Datum_LostFocus()

On Local Error GoTo LostFocusError

Select Case Index
    Case 0
        If DATE_CHECK(Format(Datum.Value, "dd/mm/yyyy"), PERIODAS_TEXT) Then
        Else
            BJPERDAT.WindowState = 0
            BJPERDAT.PeriodeBoekjaar.SetFocus
        End If
End Select
        
LostFocusError:
Exit Sub

End Sub

Private Sub FinancieelDetail_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = 45 Or KeyCode = 107 Then Volgende_Click

End Sub


Private Sub Form_Load()
    
    mfgLijst.Rows = 2
    mfgLijst.Cols = 7
    mfgLijst.Row = 0
    
    mfgLijst.Col = 0
    mfgLijst.ColAlignment(0) = flexAlignLeftCenter
    mfgLijst.text = "Lijn"
    mfgLijst.ColWidth(0) = 450
    
    mfgLijst.Col = 1
    mfgLijst.ColAlignment(1) = flexAlignLeftCenter
    mfgLijst.text = "BBA Code"
    mfgLijst.ColWidth(1) = 975
    
    mfgLijst.Col = 2
    mfgLijst.ColAlignment(2) = flexAlignLeftCenter
    mfgLijst.text = "Document"
    mfgLijst.ColWidth(2) = 1500
    
    mfgLijst.Col = 3
    mfgLijst.ColAlignment(3) = flexAlignLeftCenter
    mfgLijst.text = "TegenRek."
    mfgLijst.ColWidth(3) = 1005
        
    mfgLijst.Col = 4
    mfgLijst.text = "Bedrag"
    mfgLijst.ColWidth(4) = 990
    
    mfgLijst.Col = 5
    mfgLijst.ColAlignment(5) = flexAlignLeftCenter
    mfgLijst.text = "Omschrijving"
    mfgLijst.ColWidth(5) = 1920
            
    mfgLijst.Col = 6
    mfgLijst.text = "Cumul Saldo"
    mfgLijst.ColWidth(6) = 1215
            
    If LaadTekst("dnnInstellingen", "CodaIOMap") = "" Then
        BeWaarTekst "dnnInstellingen", "CodaIOMap", LOCATION_DESKTOP
    End If
    
    If (Right(LOCATION_COMPANYDATA, 5) = "\098\" Or Right(LOCATION_COMPANYDATA, 5) = "\099\") Then
        Me.TextBoxWarningTestCompany.Visible = True
    End If

    Dim T As Integer
    Dim TT As Integer

    Top = 0
    Left = 0

    LabelInfo(2).Caption = " Document    " + "TegenR. " + "      Bedrag " + "Omschrijving                    " + "Fin. Kort."

    Datum.Value = MIM_GLOBAL_DATE
    BeginBalans = Val(String99(READING, 64))

    ToegestaneKorting = String99(READING, 28)
    BekomenKorting = String99(READING, 27)
    DefaultRekening = String99(READING, 101)

    RecNummer(0) = 31
    RekeningNummer(0) = String99(READING, 41)
    Uittreksel(0) = LTrim$(RTrim$(String99(READING, 31)))

    RecNummer(1) = 32
    RekeningNummer(1) = String99(READING, 42)
    Uittreksel(1) = LTrim$(RTrim$(String99(READING, 32)))

    RecNummer(2) = 33
    RekeningNummer(2) = String99(READING, 43)
    Uittreksel(2) = LTrim$(RTrim$(String99(READING, 33)))

    RecNummer(3) = 34
    RekeningNummer(3) = String99(READING, 44)
    Uittreksel(3) = LTrim$(RTrim$(String99(READING, 34)))

    RecNummer(4) = 35
    RekeningNummer(4) = String99(READING, 45)
    Uittreksel(4) = LTrim$(RTrim$(String99(READING, 35)))

    RecNummer(5) = 38
    RekeningNummer(5) = String99(READING, 39)
    Uittreksel(5) = LTrim$(RTrim$(String99(READING, 38)))

    RecNummer(6) = 215
    RekeningNummer(6) = String99(READING, 211)
    Uittreksel(6) = LTrim$(RTrim$(String99(READING, 215)))

    RecNummer(7) = 216
    RekeningNummer(7) = String99(READING, 212)
    Uittreksel(7) = LTrim$(RTrim$(String99(READING, 216)))

    RecNummer(8) = 217
    RekeningNummer(8) = String99(READING, 213)
    Uittreksel(8) = LTrim$(RTrim$(String99(READING, 217)))

    RecNummer(9) = 218
    RekeningNummer(9) = String99(READING, 214)
    Uittreksel(9) = LTrim$(RTrim$(String99(READING, 218)))

    For T = 0 To 9
        If Trim(RekeningNummer(T)) = "" Then
        Else
            bGet TABLE_LEDGERACCOUNTS, 0, RekeningNummer(T)
            If Ktrl Then
                A$ = "Niet aanwezig. Installeer via Setup Boekjaar."
            Else
                RecordToVeld TABLE_LEDGERACCOUNTS
                A$ = RekeningNummer(T) + Chr(124) + RTrim$(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"))
            End If
            KeuzeInfo(0).AddItem A$
            If DefaultRekening = RekeningNummer(T) Then
                TT = T
            End If
        End If
    Next
    KeuzeInfo(0).ListIndex = TT

End Sub


Private Sub KeuzeInfo_Click(Index As Integer)

    Dim locationDesktop As String
    
    locationDesktop = LaadTekst("dnnInstellingen", "CodaIOMap")
    
    If Dir(locationDesktop + "*.xda") = "" Then
        Me.ButtonReadCamt053.Enabled = False
        Me.CheckBoxSepaViewer.Visible = False
    Else
        Me.ButtonReadCamt053.Enabled = True
        Me.CheckBoxSepaViewer.Visible = True
    End If

    Select Case Index
        Case 0
       
        Select Case UCase$(Uittreksel(KeuzeInfo(0).ListIndex))
            Case "A" To "Z"
                LabelInfo(11).Caption = Str$(Val(Right(Uittreksel(KeuzeInfo(0).ListIndex), 4)) + 1)
            Case Else
                LabelInfo(11).Caption = Str$(Val(Uittreksel(KeuzeInfo(0).ListIndex)) + 1)
        End Select
        bGet TABLE_LEDGERACCOUNTS, 0, vSet((KeuzeInfo(0).text), 7)
        If Ktrl Then
        Else
            RecordToVeld TABLE_LEDGERACCOUNTS
            If BeginBalans = 1 Then
                If bhEuro Then
                    lblInfo(0).Caption = Format(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(22 + BJPERDAT.Boekjaar.ListIndex, "000") + " #")), "#,##0.00")
                    LabelInfo(12).Caption = Format(CDbl(lblInfo(0).Caption) * EURO, "#,##0.00")
                Else
                    LabelInfo(12).Caption = Format(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(22 + BJPERDAT.Boekjaar.ListIndex, "000") + " #")), "#,##0.00")
                    lblInfo(0).Caption = Format(CDbl(LabelInfo(12)) / EURO, "#,##0.00")
                End If
            Else
                If bhEuro Then
                    lblInfo(0).Caption = Format(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(22 + BJPERDAT.Boekjaar.ListIndex, "000") + " #")) + Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#e" + Format(23 + BJPERDAT.Boekjaar.ListIndex, "000") + " #")), "#,##0.00")
                    LabelInfo(12).Caption = Format(Round(CDbl(lblInfo(0).Caption) * EURO), "#,##0.00")
                Else
                    LabelInfo(12).Caption = Format(Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(22 + BJPERDAT.Boekjaar.ListIndex, "000") + " #")) + Val(vBibTekst(TABLE_LEDGERACCOUNTS, "#v" + Format(23 + BJPERDAT.Boekjaar.ListIndex, "000") + " #")), "#,##0.00")
                    lblInfo(0).Caption = Format(CDbl(LabelInfo(12)) / EURO, "#,##0.00")
                End If
            End If
            LabelInfo(13).Caption = LabelInfo(12).Caption
            lblInfo(1).Caption = lblInfo(0).Caption
        End If
End Select

End Sub

Private Sub KeuzeInfo_LostFocus(Index As Integer)

Select Case Index
    Case 0
        If KeuzeInfo(0).text = "" Then
            Beep
            KeuzeInfo(0).SetFocus
        End If
End Select

End Sub

Private Sub LabelInfo_Change(Index As Integer)

Select Case Index
    Case 12, 13
        Select Case CDbl(LabelInfo(Index).Caption)
            Case 0
                LabelInfo(Index).BackColor = &HC0C0C0
                lblInfo(Index - 12).BackColor = &HC0C0C0
            Case Is > 0
                LabelInfo(Index).BackColor = &H80FF80
                lblInfo(Index - 12).BackColor = &H80FF80
            Case Is < 0
                LabelInfo(Index).BackColor = &H80FFFF
                lblInfo(Index - 12).BackColor = &H80FFFF
        End Select
End Select

End Sub

Private Sub Schoon()
Dim T As Integer

Do While FinancieelDetail.ListCount
    FinancieelDetail.RemoveItem 0
Loop

KeuzeInfo(0).Enabled = True
TT = KeuzeInfo(0).ListIndex
If TT = 9 Then
    KeuzeInfo(0).ListIndex = 0
Else
    KeuzeInfo(0).ListIndex = TT + 1
End If
KeuzeInfo(0).ListIndex = TT

End Sub

Private Sub Struktuur_Click()
Dim ReferteTxt As String
Dim DummyText As String
Dim dPip As Double

Dim dTotaal As Currency
Dim dBedragTekst As Currency
Dim dBedragKtrl As Currency

Dim sPip As String * 2
Dim DefaultKlanten As String * 7

SnelHelpPrint " ", BL_LOGGING
DefaultKlanten = String99(READING, 9)

      Msg = "Breng mededeling in" + vbCrLf
Msg = Msg + "met masker nnnnnnnnnnnn" + vbCrLf + vbCrLf
Msg = Msg + "Waarbij n staat voor elk" + vbCrLf
Msg = Msg + "van de 12 verplichte cijfers" + vbCrLf + vbCrLf

ReferteTxt = InputBox(Msg, "Gestruktureerde betaling")
If ReferteTxt = "" Then
    Exit Sub
Else
    dPip = Val(Left(ReferteTxt, 10))
    sPip = Format(dPip - Int(dPip / 97) * 97, "00")
    If sPip = "00" Then sPip = "97"
    If sPip <> Right(ReferteTxt, 2) Then
        MsgBox "Ongeldige invoer" + vbCrLf + vbCrLf + sPip + " <> " + Right(ReferteTxt, 2) + vbCrLf + vbCrLf + "Een gestructureerde referte heeft een kontrolesysteem.  Uw invoer is ongeldig!", 0, "Gebruikersfout"
        Exit Sub
    End If
End If

'ReferteTxt bestaat uit 12 posities
'123-4567890-12
'positie 8 moet altijd '0' zijn voor klanten

Select Case Mid(ReferteTxt, 8, 1)
    Case "0"
        bGet TABLE_CUSTOMERS, 0, Mid(ReferteTxt, 4, 4) + Mid(ReferteTxt, 9, 2)
        If Ktrl Then
            'het is geen klant, probeer als leveranciersreferte...
        Else
            RecordToVeld TABLE_CUSTOMERS
                  Msg = "Breng bedrag in voor" & vbCr
            Msg = Msg & "totaal van " & Left(ReferteTxt, 1) & " kwijtingen" & vbCr & vbCr
            Msg = Msg & "klant :" & vbCr & vbCr & RTrim$(vBibTekst(TABLE_CUSTOMERS, "#A100 #") & " " & vBibTekst(TABLE_CUSTOMERS, "#A101 #")) & " " & RTrim$(vBibTekst(TABLE_CUSTOMERS, "#A125 #") & " " & vBibTekst(TABLE_CUSTOMERS, "#A127 #")) & vbCr
            Msg = Msg & "Rekeningen:" & vBibTekst(TABLE_CUSTOMERS, "#A170 #") & " " & vBibTekst(TABLE_CUSTOMERS, "#v251 #")
            dBedragTekst = Val(InputBox(Msg, "Totaal betaling"))
            If dBedragTekst = 0 Then
                Exit Sub
            Else
            '    MsgBox "Nog in ontwerp voor fakturen.  Enkel voorlopig mogelijk voor polisbeheer"
            '    Exit Sub
            End If
        End If
    Case Else
        MsgBox "Geen R&V Gestruktureerde mededeling.", vbExclamation
        Exit Sub
End Select
bGetOrGreater TABLE_INVOICES, 1, "K" + vBibTekst(TABLE_CUSTOMERS, "#A110 #")
If Ktrl Then
    Exit Sub
Else
    RecordToVeld TABLE_INVOICES
End If

dBedragKtrl = 0
DummyText = ""
If vSet(KEY_BUF(TABLE_INVOICES), 13) <> vSet("K" + vBibTekst(TABLE_CUSTOMERS, "#A110 #"), 13) Then
    Exit Sub
Else
    GoSub LijnErBij
    Do
        bNext TABLE_INVOICES
        If Ktrl Or vSet(KEY_BUF(TABLE_INVOICES), 13) <> vSet("K" + vBibTekst(TABLE_CUSTOMERS, "#A110 #"), 13) Then
            Exit Do
        Else
            RecordToVeld TABLE_INVOICES
            GoSub LijnErBij
        End If
    Loop
    If dBedragKtrl = dBedragTekst Then
        Do While DummyText <> ""
            If Mid(DummyText, 13, 12) = Space(12) Then
            Else
                GridText = "+" + Left(DummyText, 11) + Chr(124) + DefaultKlanten + Chr(124) + Mid(DummyText, 13, 12) + Chr(124) + vSet(vBibTekst(TABLE_CUSTOMERS, "#A100 #"), 29) + Chr(124) + Space(12)
                FinancieelDetail.AddItem GridText, FinancieelDetail.ListCount
                If bhEuro Then
                    lblInfo(1).Caption = Format(CDbl(lblInfo(1).Caption) + Val(Mid(GridText, 22, 12)), "#,##0.00")
                    LabelInfo(13).Caption = Format(Round(CDbl(lblInfo(1).Caption) * EURO), "#,##0.00")
                Else
                    LabelInfo(13).Caption = Format(CDbl(LabelInfo(13).Caption) + Val(Mid(GridText, 22, 12)), "#,##0.00")
                    lblInfo(1).Caption = Format(CDbl(LabelInfo(13).Caption) / EURO, "#,##0.00")
                End If
            End If
            DummyText = Right(DummyText, Len(DummyText) - 26)
        Loop
        SnelHelpPrint GridText + " met succes bijgevoegd !", BL_LOGGING
        Exit Sub
    Else
        MsgBox "Opzoeking zonder succes, dokumentenstand : " + vbCrLf + vbCrLf + DummyText
        Exit Sub
    End If
End If
Exit Sub

LijnErBij:
If dBedragKtrl = dBedragTekst Then  'totaal al betaald
    Return
ElseIf Val(vBibTekst(TABLE_INVOICES, "#v037 #")) = Val(vBibTekst(TABLE_INVOICES, "#v249")) Then   'reeds VOLLEDIG betaald
    Return
ElseIf Left(vBibTekst(TABLE_INVOICES, "#v033 #"), 1) <> "Q" Then    'voorlopig enkel kwitanties
    Return
ElseIf dBedragTekst = Val(rsMAR(TABLE_INVOICES)("v249")) Then
    DummyText = DummyText + vBibTekst(TABLE_INVOICES, "#v033 #") + Chr(124)
    dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v249")) - Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
    DummyText = DummyText + Dec$((dTotaal), MASK_EURBH) + vbCrLf
    dBedragKtrl = dBedragKtrl + dTotaal
Else
    'MsgBox "Situatie buiten controle", vbInformation
End If
Return

End Sub


Private Sub Volgende_Click()
Dim Bedrag As Double

Unload DetailInfo
DetailInfo.Show 1
If GridText <> "" Then
    KeuzeInfo(0).Enabled = False
    FinancieelDetail.AddItem GridText, FinancieelDetail.ListCount
    Bedrag = Val(Mid(GridText, 22, 12))
    If Left(GridText, 1) = "+" Then
        If bhEuro Then
            lblInfo(1).Caption = Format(CDbl(lblInfo(1).Caption) + Bedrag, "#,##0.00")
            LabelInfo(13).Caption = Format(Round(CDbl(lblInfo(1).Caption) * EURO), "#,##0.00")
        Else
            LabelInfo(13).Caption = Format(CDbl(LabelInfo(13).Caption) + Bedrag, "#,##0.00")
            lblInfo(1).Caption = Format(CDbl(LabelInfo(13).Caption) / EURO, "#,##0.00")
        End If
    Else
        If bhEuro Then
            lblInfo(1).Caption = Format(CDbl(lblInfo(1).Caption) - Bedrag, "#,##0.00")
            LabelInfo(13).Caption = Format(Round(CDbl(lblInfo(1).Caption) * EURO), "#,##0.00")
        Else
            LabelInfo(13).Caption = Format(CDbl(LabelInfo(13).Caption) - Bedrag, "#,##0.00")
            lblInfo(1).Caption = Format(CDbl(LabelInfo(13).Caption) / EURO, "#,##0.00")
        End If
    End If
End If

End Sub

Private Function WegBoekFout() As Integer
Dim T               As Integer
Dim TotaalBedrag    As Double
Dim FinKort         As Double

Dim BeginSaldo As Currency
Dim EindSaldo As Currency
Dim VerschilSaldo As Currency

Screen.MousePointer = vbHourglass
WegBoekFout = True
DKTRL_CUMUL = 0: DKTRL_BEF = 0: DKTRL_EUR = 0
Unload frmBoeking
frmBoeking.Hide

TLB_RECORD(TABLE_JOURNAL) = ""
vBib TABLE_JOURNAL, Left(KeuzeInfo(0).text, 7), "v019"
vBib TABLE_JOURNAL, Format(Datum.Value, "yyyymmdd"), "v066"
vBib TABLE_JOURNAL, Format(Datum.Value, "yyyymmdd"), "v035"

bGet TABLE_LEDGERACCOUNTS, 0, Left(KeuzeInfo(0).text, 7)
If Ktrl Then
    Exit Function
Else
    RecordToVeld TABLE_LEDGERACCOUNTS
    dokumentSleutel = UCase$(Left(vBibTekst(TABLE_LEDGERACCOUNTS, "#v020 #"), 2)) + Right(Format(Datum.Value, "dd/mm/yyyy"), 2) + Format(Val(LabelInfo(11).Caption), "0000")
End If
vBib TABLE_JOURNAL, dokumentSleutel, "v038"
If bhEuro Then
    BeginSaldo = CDbl(lblInfo(0).Caption)
    EindSaldo = CDbl(lblInfo(1).Caption)
    VerschilSaldo = EindSaldo - BeginSaldo
    vBib TABLE_JOURNAL, "Sld:" + Dec(CDbl(lblInfo(0).Caption), MASK_EURBH) + "/" + Dec(CDbl(lblInfo(1).Caption), MASK_EURBH), "v067"
    vBib TABLE_JOURNAL, Dec((VerschilSaldo), MASK_SY(2)), "v068"
Else
    vBib TABLE_JOURNAL, "Sld : " + Dec(CDbl(LabelInfo(12).Caption), MASK_SY(0)) + " - " + Dec(CDbl(LabelInfo(13).Caption), MASK_SY(0)), "v067"
    vBib TABLE_JOURNAL, Dec$(CDbl(LabelInfo(13).Caption) - CDbl(LabelInfo(12).Caption), MASK_SY(0)), "v068"
End If
bInsert TABLE_JOURNAL, 2
If Ktrl Then Exit Function

vBib TABLE_JOURNAL, Left(KeuzeInfo(0).text, 7), "v069"
For T = 1 To FinancieelDetail.ListCount
    FinancieelDetail.ListIndex = T - 1
    If Left(FinancieelDetail.text, 1) = "+" Then
        Ontvangst = True
    Else
        Ontvangst = False
    End If
    If Mid(FinancieelDetail.text, 2, 11) = Space$(11) Then
        vBib TABLE_JOURNAL, " ", "v033"
        vBib TABLE_JOURNAL, " ", "v034"
    Else
        bGet TABLE_INVOICES, 0, Mid(FinancieelDetail.text, 2, 11)
        If Ktrl Then
            Exit Function
        Else
            RecordToVeld TABLE_INVOICES
            vBib TABLE_JOURNAL, vBibTekst(TABLE_INVOICES, "#v033 #"), "v033"
            vBib TABLE_JOURNAL, vBibTekst(TABLE_INVOICES, "#v034 #"), "v034"
        End If
    End If
    
    TotaalBedrag = Val(Mid(FinancieelDetail.text, 22, 12))
    vBib TABLE_JOURNAL, Trim$(Mid(FinancieelDetail.text, 35, 29)), "v067"
    If Val(Mid(FinancieelDetail.text, 65, 12)) <> 0 Then
        FinKort = Val(Mid(FinancieelDetail.text, 65, 12))
        TotaalBedrag = TotaalBedrag + FinKort
        Select Case Ontvangst
            Case True
                vBib TABLE_JOURNAL, Str$(FinKort), "v068"
                vBib TABLE_JOURNAL, ToegestaneKorting, "v019"
            Case False
                vBib TABLE_JOURNAL, Str$(-FinKort), "v068"
                vBib TABLE_JOURNAL, BekomenKorting, "v019"
            Case Else
                MsgBox "Stop"
                Exit Function
        End Select
        bInsert TABLE_JOURNAL, 2
        If Ktrl Then Exit Function
    End If
    
    vBib TABLE_JOURNAL, Mid(FinancieelDetail.text, 14, 7), "v019"
    Select Case Ontvangst
        Case True
            vBib TABLE_JOURNAL, Str$(-TotaalBedrag), "v068"
        Case False
            vBib TABLE_JOURNAL, Str$(TotaalBedrag), "v068"
        Case Else
            MsgBox "Stop"
            Exit Function
    End Select
    bInsert TABLE_JOURNAL, 2
    If Ktrl Then Exit Function
    If LTrim$(vBibTekst(TABLE_JOURNAL, "#v033 #")) = "" Then
    Else
        If XisEuroWisBEF = False Then
            vBib TABLE_INVOICES, Str$(Val(vBibTekst(TABLE_INVOICES, "#v037 #")) + TotaalBedrag), "v037"
        Else
            vBib TABLE_INVOICES, Str$(Val(vBibTekst(TABLE_INVOICES, "#v037 #")) + Round(TotaalBedrag / EURO, 2)), "v037"
        End If
        vBib TABLE_INVOICES, vBibTekst(TABLE_JOURNAL, "#v038 #"), "v038"
        bUpdate TABLE_INVOICES, 0
        If Ktrl Then Exit Function
    End If
Next

Screen.MousePointer = vbNormal
If DKTRL_CUMUL <> 0 Then
    frmBoeking.cmdBoeken.Enabled = False
    MsgBox "LogikaFout bij vierkantskontrole journaal." + vbCrLf + vbCrLf + "Deze verrichting wordt geannuleerd.  Controleer zelf eerst en/of raadpleeg ons."
    frmBoeking.Show 1
    WegBoekFout = True
ElseIf JournaalLocked = True Then
    frmBoeking.cmdBoeken.Enabled = False
    frmBoeking.Show 1
    WegBoekFout = True
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
    If DKTRL_CUMUL = 0 Then WegBoekFout = False
End If

End Function



Function fnBeginOpname(DeString As String) As Boolean
    
    sDatumAanmaak = Mid(DeString, 6, 6)
    sToepassingsCode = Mid(DeString, 15, 2)
    sNaamBestemmeling = Mid(DeString, 35, 26)
    If sToepassingsCode = "05" Then
        fnBeginOpname = True
    Else
        fnBeginOpname = False
    End If

End Function


Function fnOudSaldo(DeString As String) As Boolean

    sRekeningNummer = Mid(DeString, 6, 16)
    sUittreksel = Mid(DeString, 3, 3)   '?
    cOudSaldo = Val(Mid(DeString, 44, 15)) / 1000
    sDatumOudSaldo = Mid(DeString, 59, 6)
    sNaamRekeninghouder = Mid(DeString, 65, 26)
    sOmschrijvingRekening = Mid(DeString, 91, 35)
    fnOudSaldo = True
    iOptelControle = iOptelControle + 1
    
End Function


Function fnNieuwSaldo(DeString As String) As Boolean

    sRekeningNummer2 = Mid(DeString, 5, 12)
    sUittreksel2 = Mid(DeString, 2, 3)   '?
    cNieuwSaldo = Val(Mid(DeString, 43, 15)) / 1000
    sDatumNieuwSaldo = Mid(DeString, 58, 6)
    fnNieuwSaldo = True
    iOptelControle = iOptelControle + 1
    
End Function


Function fnEindOpname(DeString As String) As Boolean

    iOptelCtrlCheckUp = Val(Mid(DeString, 17, 6))
    If iOptelCtrlCheckUp <> iOptelControle Then
        SnelHelpPrint "Onlogische situatie", False
    End If
    cDebetSaldo = Val(Mid(DeString, 23, 15)) / 1000
    cCreditSaldo = Val(Mid(DeString, 38, 15)) / 1000
    If Mid(DeString, 128, 1) = "1" Then
        fnEindOpname = False
    Else
        fnEindOpname = True
    End If
    
End Function

Function fnBeweging(DeString As String) As Boolean

    iOptelControle = iOptelControle + 1
    Select Case Mid(DeString, 2, 1)
        Case "1" 'Deel 1
            sRefFinInstelling = Mid(DeString, 11, 21)
            sDC = Mid(DeString, 32, 1) '"0" = credit "1" = debet
            cBedrag = Val(Mid(DeString, 33, 15)) / 1000
            sValutadatum = Mid(DeString, 48, 6)
            
            '01 tot 39 : binnenlandse-lokale/SEPA verrichtingen 41 tot 79 : buitenlandse/non-SEPA verrichtingen 80 tot 89 : andere families
            sVerrichting = Mid(DeString, 54, 8) '01xx OVERSCHRIJVINGEN (C)
                                                '0301 CHEQUE UW BETALING (D)
                                                '03xx CHEQUE ONTVANGEN (C)
                                                '0503 ONBETAALDE INVORDERING (D)
                                                '0556 ONUITVOERBARE TERUGBETALING (C)
                                                '09xx STORTINGEN (C)
            sMededeling = Mid(DeString, 62, 1)  '"1" = gestructureerd
            If sMededeling <> "1" Then
                'verwittiging nog geven voor manuele verwerking
            End If
            sMDDZone1 = Mid(DeString, 63, 3)    '
            sMDDZone2 = Mid(DeString, 66, 50)
            sBoekDatum = Mid(DeString, 116, 6)
            sDagAfschriftVolgNummer = Mid(DeString, 122, 3)
    
        Case "2" 'Deel 2
            sMededeling2 = Mid(DeString, 11, 53)
            MsgBox "todo"
            'For COUNT_TO = 0 To 1
            '    sRefKlant(COUNT_TO) = Mid(DeString, 64 + (COUNT_TO * 13), 13)
            'Next
            'sMuntVerrichting = Mid(DeString, 90, 3)
            'cBedragMunt = Val(Mid(DeString, 93, 15)) / 1000
        
        Case "3" 'Deel 3
            sRekeningTP = Mid(DeString, 11, 37)
            MsgBox "todo"
            'sITcodesTP = Mid(DeString, 23, 10)
            'sRekeningTPextra = Mid(DeString, 33, 15)
            
            sNaamEnAdres(0) = Mid(DeString, 48, 35)
            
        
        Case Else
            MsgBox "Onlogische situatie", vbCritical
    
    End Select
    If Mid(DeString, 126, 1) = "0" Then 'er is géén vervolg !
        fnBeweging = False
    Else
        fnBeweging = True
    End If
    
End Function

Function BewegingSchoon()

    sRefFinInstelling = ""
    sRefFinInstelling2 = ""
    sDC = ""
    cBedrag = 0
    sVerrichting = ""
    sMededeling = ""
    sMDDZone1 = ""
    sMDDZone2 = ""
    sBoekDatum = ""
    sDagAfschriftVolgNummer = ""
    sValutadatum = ""
    sMededeling2 = ""
    sRefKlant(0) = ""
    sRefKlant(1) = ""
    sMuntVerrichting = ""
    cBedragMunt = 0
    sRekeningTP = ""
    sITcodesTP = ""
    sRekeningTPextra = ""
    sNaamEnAdres(0) = ""
    sNaamEnAdres(1) = ""
    sNaamEnAdres(2) = ""
    
End Function

Private Sub xdaParseToArray(inputText As String)
    
    Dim lines() As String
    Dim fields() As String
    Dim result() As String
    Dim i As Long, j As Long
    Dim maxCols As Long
    
    On Local Error GoTo 0
    
    ' Split into lines
    lines = Split(inputText, vbCrLf)
    
    ' Find maximum number of columns
    For i = LBound(lines) To UBound(lines)
        fields = Split(lines(i), vbTab)
        If UBound(fields) + 1 > maxCols Then
            maxCols = UBound(fields) + 1
        End If
    Next i
    
    ' Dimension result array: rows = number of lines, cols = max fields
    ReDim xdaLinesDATAArray(UBound(lines), maxCols - 1)
    
    ' Fill the array
    For i = LBound(lines) To UBound(lines)
        fields = Split(lines(i), vbTab)
        For j = LBound(fields) To UBound(fields)
            xdaLinesDATAArray(i, j) = fields(j)
        Next j
    Next i
    
End Sub

