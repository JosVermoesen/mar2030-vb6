VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form JaarBtwBelgie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jaarlisting BTW België"
   ClientHeight    =   1560
   ClientLeft      =   540
   ClientTop       =   4200
   ClientWidth     =   7320
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
   HelpContextID   =   50700
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1560
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox tbRefDeclarant 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2760
      TabIndex        =   2
      Top             =   1080
      Width           =   2745
   End
   Begin VB.TextBox tbRefRepresentative 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "04400582170091"
      Top             =   1080
      Width           =   2505
   End
   Begin VB.CommandButton CmdEmailNBB 
      Caption         =   "Genereren en verzenden"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5640
      Picture         =   "0600014.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1545
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
      Height          =   375
      Left            =   5640
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1545
   End
   Begin MSComCtl2.DTPicker Datum 
      Height          =   315
      Left            =   960
      TabIndex        =   0
      Top             =   120
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
      Format          =   73990147
      CurrentDate     =   39083
      MaxDate         =   73050
      MinDate         =   31048
   End
   Begin MSMAPI.MAPISession MPISessie 
      Left            =   2760
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSMAPI.MAPIMessages MPIBericht 
      Left            =   4320
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Referte Aangever"
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
      Index           =   3
      Left            =   2760
      TabIndex        =   6
      Top             =   720
      Width           =   2115
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Referte Vertegenwoordiger"
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
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   2115
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datu&m"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "JaarBtwBelgie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim psTekst(5) As String        'Koptekstinfo
Dim aa As String

Dim RapportSelektie As String
Dim RapportTitel As String
Dim RapportDefinitie As String
Dim LijstNaam As String
Dim VeldTXT(20) As String

Dim dTotaal As Currency
Dim dBeTeWe As Currency

Dim dTOT As Currency
Dim dBTW As Currency

Dim VsoftVanaf          As String
Dim VsoftTot            As String

Dim ktrlEur As Boolean
Dim desktopLocatie As String

Private Sub Annuleren_Click()

Unload Me

End Sub

Private Sub CmdEmailNBB_Click()

Dim BeginSleutel As String
Dim EindSleutel As String
Dim KopBuf As String
Dim XX As Integer
Dim Teller As Integer
Dim FlTabBestand As Integer

Dim pathString As String * 260
Dim RetInteger As Integer

Dim dTnt1 As Currency
Dim dTnt2 As Currency
Dim Lijn As Integer
Dim BedragMinstens As Currency

dTotaal = 0
dBeTeWe = 0
BedragMinstens = 250


BeginSleutel = vSet("K0", 13)
EindSleutel = vSet("K" + "zzzzzzzzzzzz", 13)

    InitVelden

ktrlEur = EurBefJaar()

bGetOrGreater TABLE_INVOICES, 1, BeginSleutel
If Ktrl Or UCase(KEY_BUF(TABLE_INVOICES)) > UCase(EindSleutel) Then
    Beep
    Exit Sub
ElseIf Left(KEY_BUF(TABLE_INVOICES), 1) <> Left(BeginSleutel, 1) Then
    Beep
    Exit Sub
Else
    VeldTXT(0) = Chr$(0)
    Screen.MousePointer = vbHourglass
    
    
    PrintTitel
    RecordToVeld TABLE_INVOICES
    GoSub KontroleVoorwaarden
End If

Me.Refresh

Do
    bNext TABLE_INVOICES
    If Ktrl Or UCase(KEY_BUF(TABLE_INVOICES)) > UCase(EindSleutel) Then
        Exit Do
    End If
    RecordToVeld TABLE_INVOICES
    GoSub KontroleVoorwaarden
Loop
    
        VeldTXT(7) = Dec$((dTnt1), MASK_EUR)
        VeldTXT(8) = Dec$((dTnt2), MASK_EUR)
    GoSub PrintDeLijn
    EindTotaal
    Screen.MousePointer = vbNormal
    GoSub BNvraag
    Annuleren_Click
Exit Sub

KontroleVoorwaarden:
'Enkel het reële jaar
If Left(vBibTekst(TABLE_INVOICES, "#v035 #"), 4) <> Right(Format(Datum, "dd/mm/yyyy"), 4) Then
    Return
End If

dTOT = 0
Select Case Left(vBibTekst(TABLE_INVOICES, "#v033 #"), 1)
    Case "V"
        For Teller = 55 To 63
            dTOT = dTOT + Val(vBibTekst(TABLE_INVOICES, "#v" + Format(Teller, "000") + " #"))
        Next
        dTOT = dTOT + Val(vBibTekst(TABLE_INVOICES, "#v089 #"))
        dBTW = Val(vBibTekst(TABLE_INVOICES, "#v064 #"))
        If Mid(vBibTekst(TABLE_INVOICES, "#v033 #"), 2, 1) = "1" Then
            dTOT = -dTOT
            dBTW = -dBTW
        End If
        If ktrlEur = True Then
            If vBibTekst(TABLE_INVOICES, "#v035 #") > Left(BOOKYEAR_FROMTO, 8) Then
            Else
                dTOT = Round(dTOT / EURO, 2)
                dBTW = Round(dBTW / EURO, 2)
            End If
        End If
        
    Case Else
        Return
End Select

If Trim$(vBibTekst(TABLE_INVOICES, "#v034 #")) <> Trim$(KopBuf) Then
    If VeldTXT(0) = Chr$(0) Then
    Else
            VeldTXT(7) = Dec$((dTnt1), MASK_EUR)
            VeldTXT(8) = Dec$((dTnt2), MASK_EUR)
        
        GoSub PrintDeLijn
    End If
    
    dTnt1 = 0
    dTnt2 = 0
    bGet TABLE_CUSTOMERS, 0, vSet(Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2, 12), 12)
    KopBuf = vBibTekst(TABLE_INVOICES, "#v034 #")
    VeldTXT(0) = Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2, 12)
    If Ktrl Then
        VeldTXT(1) = "* niet meer aanwezig *"
    Else
        RecordToVeld TABLE_CUSTOMERS
        VeldTXT(1) = vSet(vBibTekst(TABLE_CUSTOMERS, "#A100 #"), 32)
        SnelHelpPrint VeldTXT(1), BL_LOGGING
        VeldTXT(2) = vSet(RTrim$(RTrim$(vBibTekst(TABLE_CUSTOMERS, "#A104 #")) + " " + RTrim$(vBibTekst(aIndex, "#A105 #")) + " " + RTrim$(vBibTekst(aIndex, "#A106 #"))), 24)
        VeldTXT(3) = vSet(vBibTekst(TABLE_CUSTOMERS, "#A107 #"), 4)
        VeldTXT(4) = vSet(RTrim$(vBibTekst(TABLE_CUSTOMERS, "#A108 #")), 23)
        VeldTXT(5) = "BE"
        VeldTXT(6) = vSet(vBibTekst(TABLE_CUSTOMERS, "#A161 #"), 10)
    End If
End If

    dTnt1 = dTnt1 + dTOT
    dTnt2 = dTnt2 + dBTW
Return

PrintDeLijn:
If dTnt1 < BedragMinstens Then
    Return
ElseIf vBibTekst(TABLE_CUSTOMERS, "#v150 #") <> "BE" Then
    Return
ElseIf RTrim$(vBibTekst(TABLE_CUSTOMERS, "#A161 #")) = "" Then
    Return
Else
    If RTrim$(vBibTekst(TABLE_CUSTOMERS, "#A161 #")) = BtwKontrole(RTrim$(vBibTekst(TABLE_CUSTOMERS, "#A161 #")), False) Then
    Else
        Msg = "Het BTW nummer : " + RTrim$(vBibTekst(TABLE_CUSTOMERS, "#A161 #")) + " is foutief !" + vbCrLf
        Msg = Msg + "Klant : " + vBibTekst(TABLE_CUSTOMERS, "#A110 #") + " " + vBibTekst(TABLE_CUSTOMERS, "#A100 #") + vbCrLf + vbCrLf
        Msg = Msg + "Een 0000000000 nummer wordt ingevoegd.  Gelieve de kontrolelijst" + vbCrLf
        Msg = Msg + "na te kijken én te verbeteren a.u.b. vooraleer door te sturen..."
        MsgBox Msg
        vBib TABLE_CUSTOMERS, "0000000000", "A161 "
        VeldTXT(6) = vSet(vBibTekst(TABLE_CUSTOMERS, "#A161 #"), 10)
        bUpdate TABLE_CUSTOMERS, 0
    End If
    'kontrole btw nummer
    dTotaal = dTotaal + dTnt1
    dBeTeWe = dBeTeWe + dTnt2
    Lijn = Lijn + 1
    VeldTXT(0) = Dec$((Lijn), "000000")
    PrintVelden
    VeldTXT(0) = ""
End If
Return

BNvraag:
Msg = "Positionele jaarlisting op uw Bureaublad doormailen naar Vsoft voor afhandeling" & vbCrLf & vbCrLf & "Kies 'ja' voor doormailen (aanbevolen), 'nee' indien U zelf afhandelt."
Ktrl = MsgBox(Msg, vbYesNoCancel + vbDefaultButton3 + vbQuestion)
Select Case Ktrl
    Case 2
        Msg = "Taak verlaten zonder enige verwerking.  Bent U zeker"
        Ktrl = MsgBox(Msg, vbYesNo + vbDefaultButton2 + vbQuestion)
        If Ktrl = 7 Then
            GoTo BNvraag
        Else
            Exit Sub
        End If
        
    Case 6

        Screen.MousePointer = vbHourglass
        If Me.MPISessie.SessionID <> 0 Then
        Else
            On Local Error Resume Next
            Me.MPISessie.SignOn
            If Err Then
                SnelHelpPrint Err.Description, BL_LOGGING
            Else
                On Local Error GoTo 0
                Me.MPIBericht.SessionID = Me.MPISessie.SessionID
                SnelHelpPrint "E-mail sessie met succes opgestart. IDkode :" + Format(Me.MPISessie.SessionID), BL_LOGGING
            End If
        End If
        On Local Error GoTo MPIError
        'Compose New Message
        Me.MPIBericht.Compose
                        
        'Address message
        Me.MPIBericht.RecipDisplayName = "vsoftINTERVAT"
        Me.MPIBericht.RecipAddress = "SMTP:info@rv.be"
            
        'Resolve recipient name
        Me.MPIBericht.AddressResolveUI = True '=dialogbox, false = error genereren
                    
        'Create the message
        Me.MPIBericht.MsgSubject = "Btwjaarlisting"
        Me.MPIBericht.MsgNoteText = "Formaat:Positioneel Bericht.  Aangemaakt met marIntegraal versie " & MAR_VERSION
        Me.MPIBericht.AttachmentPathName = desktopLocatie & "\" + Me.tbRefDeclarant + ".xml"
            
        'Send the message
        On Error Resume Next
        Me.MPIBericht.Send True
        Screen.MousePointer = vbNormal
        If Err = 32001 Then
            MsgBox "U koos om de mail af te breken.", vbInformation
            GoTo BNvraag
        Else
            MsgBox "Zorg ervoor dat uw mailtoepassing effectief kan verzenden nu of straks.  U ontvangt later nog ontvangstbevestiging vanwege onze diensten", vbInformation
            Exit Sub
        End If
        
    Case 7
        'niks
        'Zelf afhandelen (enkel nog afdruk afleveringnota's hierna
        Msg = "Het bestand " & Me.tbRefDeclarant & ".xml" & " staat op het bureaublad ter beschikking." & vbCrLf & vbCrLf & "Volg de instructies in uw intervat software of volg instructies van uw raadgever."
        MsgBox Msg, vbInformation
End Select

Return

MPIError:
MsgBox Error
Screen.MousePointer = vbNormal
Resume Next

End Sub

Private Sub Datum_Change()

JaarBtwBelgie.Caption = "Jaarlisting B.T.W. België " + Right(Format(Datum, "dd/mm/yyyy"), 4) + " !!!"
LijstNaam = JaarBtwBelgie.Caption

End Sub


Private Sub EindTotaal()
Dim T As Integer

For T = 0 To 8
    VeldTXT(T) = ""
Next
VeldTXT(0) = "Totaal"
    VeldTXT(7) = Dec$((dTotaal), MASK_EUR)
    VeldTXT(8) = Dec$((dBeTeWe), MASK_EUR)

T = 0
aa = ""
Do While REPORT_TAB(T) <> 0
    aa = aa & VeldTXT(T) & vbTab
    T = T + 1
Loop


    Xlog.X.AddItem aa, Xlog.X.Rows - 1
    Xlog.X.ColWidth(0) = 705
    Xlog.X.ColWidth(1) = 2055
    Xlog.X.ColWidth(2) = 2130
    Xlog.X.ColWidth(3) = 495
    Xlog.X.ColWidth(4) = 1485
    Xlog.X.ColWidth(5) = 315
    Xlog.X.ColWidth(6) = 1065
    Xlog.X.ColWidth(7) = 900
    Xlog.X.ColWidth(8) = 900
    Xlog.WindowState = 2

    Xlog.X.ColAlignment(7) = flexAlignRightTop
    Xlog.X.ColAlignment(8) = flexAlignRightTop
    
    'XLog.X.GridLines = False
    Screen.MousePointer = vbNormal
    Xlog.Afsluiten.Enabled = False
    Xlog.WijzigenLijn.Visible = False
    Xlog.Afsluiten.TabStop = False
    Xlog.cbAfbeelding.Visible = False
    XLogKey = ""
    Xlog.SSTab1.TabVisible(1) = False
    Xlog.Show 1
    Ktrl = XmlGenerate
    Unload Xlog

Exit Sub

PrtHandler4:
MsgBox "Kontroleer de printer."
Resume

End Sub

Function XmlGenerate() As Boolean

Dim T As Long

Dim totalAmount As String
Dim totalVat As String

'get totalAmount
Xlog.X.Row = Xlog.X.Rows - 2
Xlog.X.Col = 7
totalAmount = Trim(Xlog.X.Clip)

'get totalVat
Xlog.X.Col = 8
totalVat = Trim(Xlog.X.Clip)

Dim yearListingMain As String
Dim vatYearClient As String
Dim clientLine As String
Dim listClientLines As String

Ktrl = ScrLeesBestandAlleTekst(yearListingMain, PROGRAM_LOCATION + "xml-templates\be-vatyearlisting-main.xml")
If Ktrl = 0 Then
    MsgBox "Onverwachte situatie", vbCritical
End If

'{clientsCount}
yearListingMain = Replace(yearListingMain, "{clientsCount}", Trim(Str(Xlog.X.Rows - 3)))
'totals {turnoverSum} and {vatAmountSum}
Xlog.X.Row = Xlog.X.Rows - 2
Xlog.X.Col = 7
yearListingMain = Replace(yearListingMain, "{turnoverSum}", Trim(Xlog.X.Clip))
Xlog.X.Col = 8
yearListingMain = Replace(yearListingMain, "{vatAmountSum}", Trim(Xlog.X.Clip))
'{referenceDeclarant}
yearListingMain = Replace(yearListingMain, "{referenceDeclarant}", Trim(Me.tbRefDeclarant.Text))
'{fiscalYear}
yearListingMain = Replace(yearListingMain, "{fiscalYear}", Right(Datum, 4))
yearListingMain = Replace(yearListingMain, "{vatNumberDeclarant}", String99(READING, 51))

Dim tmpString As String

tmpString = String99(READING, 46)
If InStr(tmpString, "&") Then 'verbeteren voor XML bestand!!!
    tmpString = Replace(tmpString, "&", "&amp;")
End If
yearListingMain = Replace(yearListingMain, "{nameDeclarant}", Trim(tmpString))
yearListingMain = Replace(yearListingMain, "{streetDeclarant}", Trim(String99(READING, 47)))
yearListingMain = Replace(yearListingMain, "{postalCodeDeclarant}", Mid(String99(READING, 48), 1, 4))
yearListingMain = Replace(yearListingMain, "{cityDeclarant}", Mid(String99(READING, 48), 6))
yearListingMain = Replace(yearListingMain, "{emailDeclarant}", Trim(String99(READING, 50)))

Ktrl = ScrLeesBestandAlleTekst(vatYearClient, PROGRAM_LOCATION + "xml-templates\be-vatyearclient.xml")
If Ktrl = 0 Then
    MsgBox "Onverwachte situatie", vbCritical
End If

For T = 1 To Xlog.X.Rows - 3
    If T > 1 Then
        listClientLines = listClientLines + vbCrLf
    End If
    clientLine = vatYearClient
    Xlog.X.Row = T
    clientLine = Replace(clientLine, "{sequenceNumber}", Trim(Str(T)))
    Xlog.X.Col = 6
    clientLine = Replace(clientLine, "{companyVatNumber}", Trim(Xlog.X.Clip))
    Xlog.X.Col = 7
    clientLine = Replace(clientLine, "{turnOver}", Trim(Xlog.X.Clip))
    Xlog.X.Col = 8
    clientLine = Replace(clientLine, "{vatAmount}", Trim(Xlog.X.Clip))
    listClientLines = listClientLines + clientLine
Next
yearListingMain = Replace(yearListingMain, "<Vsoft>vatYearClientsList</Vsoft>", listClientLines)
Ktrl = ScrMaakTekstBestand(yearListingMain, LOCATION_COMPANYDATA + "vat\" + Me.tbRefDeclarant + ".xml")
If Ktrl = 0 Then
    MsgBox "Fout bij bewaren van " + Me.tbRefDeclarant + ".xml", vbCritical
End If

Dim Ret As Integer
Dim path As String * 260

'Application Data lokatie ?
Ret = SHGetFolderPath(0, 0, 0, 0, path)
desktopLocatie = Left(path, InStr(path, Chr(0)) - 1)
Ktrl = ScrMaakTekstBestand(yearListingMain, desktopLocatie + "\" + Me.tbRefDeclarant + ".xml")
If Ktrl = 0 Then
    MsgBox "Fout bij bewaren van " + Me.tbRefDeclarant + ".xml", vbCritical
End If

End Function

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

If String99(READING, 51) = BtwKontrole(String99(READING, 51), False) Then
Else
    MsgBox "ongeldig btw nummer in setup", vbCritical
    Annuleren_Click
    Exit Sub
End If


Dim TempB_MODUS As Integer

For COUNT_TO = 0 To BJPERDAT.PeriodeBoekjaar.ListCount
    If Mid(BJPERDAT.PeriodeBoekjaar.List(COUNT_TO), 14, 5) = "31/12" Then
        Datum.Value = Mid(BJPERDAT.PeriodeBoekjaar.List(COUNT_TO), 14)
        Datum.Enabled = False
        Exit For
    End If
Next

JaarBtwBelgie.Caption = "Jaarlisting B.T.W. België " + Right(Format(Datum, "dd/mm/yyyy"), 4) + " !!!"
LijstNaam = JaarBtwBelgie.Caption
TopDown.Caption = LijstNaam
Me.tbRefDeclarant.Text = "L" + String99(READING, 51) + Right(Format(Datum, "dd/mm/yyyy"), 4)

End Sub


Private Sub InitVelden()
Dim T As Integer

REPORT_FIELD(0) = "Lijn"
    REPORT_TAB(0) = 2

REPORT_FIELD(1) = "Naam"
    REPORT_TAB(1) = 9
    
REPORT_FIELD(2) = "Straat en Nummer"
    REPORT_TAB(2) = 42

REPORT_FIELD(3) = "P/C"
    REPORT_TAB(3) = 67

REPORT_FIELD(4) = "Stad"
    REPORT_TAB(4) = 72

REPORT_FIELD(5) = "LN"
    REPORT_TAB(5) = 96

REPORT_FIELD(6) = "BTW Nr."
    REPORT_TAB(6) = 99

REPORT_FIELD(7) = "Goederen"
    REPORT_TAB(7) = 109

REPORT_FIELD(8) = "BTW"
    REPORT_TAB(8) = 119

REPORT_TAB(9) = 0

    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = "Jaarlisting BTW België"
    Xlog.X.Cols = 9
    Xlog.X.Row = 0
    For T = 0 To 8
        Xlog.X.Col = T
        Xlog.X.Text = REPORT_FIELD(T)
    Next
    Me.Show

End Sub

Private Sub PrintTitel()
Dim T As Integer


End Sub

Private Sub PrintUserDef(TypeEnTaal As String)
Dim Dummy As String
Dim BeginXbox As Integer
Dim BeginYbox As Integer
Dim T               As Integer
Dim Tabulatie       As Integer
Dim FlFree          As Integer
Dim Teller          As Integer

ReDim PPsTekst(50) As String
ReDim psX(50) As Single
ReDim psY(50) As Single
ReDim psFontSize(50) As Single
ReDim psFontName(50) As String
ReDim psFontBold(50) As Integer
ReDim psFontItalic(50) As Integer
ReDim psFontUnderLine(50) As Integer
ReDim psColor(50) As Long
Dim MaxPslokatie        As Integer

Dim BoxTeller As Integer
ReDim Box(10) As Integer
ReDim BeginX(10) As Integer
ReDim BeginY(10) As Integer
ReDim TotX(10) As Integer
ReDim TotY(10) As Integer
Dim PsLokatie As Integer

On Local Error GoTo ErrorLoad

If Dir$(LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + ".Txt") = "" Then
    Beep
    Exit Sub
Else
    FlFree = FreeFile
    Open LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + ".Txt" For Input As FlFree
    Input #FlFree, Dummy
    Input #FlFree, VsoftVanaf
    Input #FlFree, VsoftTot
    For Teller = 0 To 10
        Input #FlFree, BeginX(Teller), BeginY(Teller), TotX(Teller), TotY(Teller), Box(Teller)
    Next
    Teller = 0
    While Not EOF(FlFree)
        Input #FlFree, PPsTekst(Teller), psX(Teller), psY(Teller), psFontSize(Teller), psFontName(Teller), psFontBold(Teller), psFontItalic(Teller), psFontUnderLine(Teller), psColor(Teller)
        Teller = Teller + 1
    Wend
    MaxPslokatie = Teller
    Close FlFree
End If

BeginXbox = BeginX(0)
BeginYbox = BeginY(0)

PsLokatie = 0
Do While PsLokatie <= MaxPslokatie
    If PPsTekst(PsLokatie) <> "" Then
        Printer.FontName = psFontName(PsLokatie)
        Printer.FontItalic = psFontItalic(PsLokatie)
        Printer.FontSize = psFontSize(PsLokatie)
        Printer.FontBold = psFontBold(PsLokatie)
        Printer.FontUnderline = psFontUnderLine(PsLokatie)
        Printer.ForeColor = psColor(PsLokatie)
    End If
    Printer.CurrentX = psX(PsLokatie)
    Printer.CurrentY = psY(PsLokatie)
    Printer.Print PPsTekst(PsLokatie);
    PsLokatie = PsLokatie + 1
Loop
For T = 0 To 10
    If BeginX(T) = 0 Then
    ElseIf Box(T) Then
        Printer.Line (BeginX(T), BeginY(T))-(TotX(T), TotY(T)), , B
    Else
        'printer.Line (BeginX(T), BeginY(T))-(TotX(T), TotY(T))
    End If
Next

Dim FiguurX As Integer
Dim FiguurY As Integer
Dim FiguurName As String
Dim FigBestandsnaam As String
Dim Hoogte As Variant
Dim Breedte As Variant

If Dir$(LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + "G.Txt") = "" Then
Else
    FlFree = FreeFile
    Open LOCATION_COMPANYDATA + "DDEF" + TypeEnTaal + "G.Txt" For Input As FlFree
    While Not EOF(FlFree)
        Input #FlFree, FiguurX, FiguurY, FiguurName
        If ScrLeesTekstBestand(Msg, LOCATION_COMPANYDATA & FiguurName & ".mfd") Then
            FigBestandsnaam = Left(Msg, InStr(Msg, vbTab) - 1)
            Mim.imgFiguur.Picture = LoadPicture(FigBestandsnaam)
            Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
            Hoogte = Left(Msg, InStr(Msg, vbTab) - 1)
            Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
            Breedte = Msg
            Printer.PaintPicture Mim.imgFiguur.Picture, FiguurX, FiguurY, Breedte, Hoogte
        Else
            MsgBox "afdrukprobleemke figuurobject: " & FiguurName
        End If
    Wend
    Close FlFree
End If
Exit Sub

ErrorLoad:
MsgBox "Stop tijdens inladen dokumentdefinitie" + vbCrLf + vbCrLf + Error$
Exit Sub
Resume

End Sub

Private Sub PrintVelden()
Dim T As Integer
Dim VeldTekst As String

T = 0
aa = ""
Do While REPORT_TAB(T) <> 0
    aa = aa & VeldTXT(T) & vbTab
    T = T + 1
Loop
Xlog.X.AddItem aa, Xlog.X.Rows - 1

End Sub

Private Sub tbRefDeclarant_Change()

    MsgBox "Wijzigingen aanbrengen wordt niet aanbevolen.", vbOKOnly + vbExclamation
    
End Sub

Private Sub tbRefRepresentative_GotFocus()

MsgBox "Wijzigingen worden hier genegeerd." & vbCrLf & vbCrLf & "Wenst U te verzenden via een andere dienstverlener, gelieve het xml bestand aan deze over te maken", vbOKOnly + vbInformation, "Vertegenwoordiger 0440058217"
tbRefRepresentative = "04400582170091"

End Sub


Function EurBefJaar() As Boolean

Dim tempoBoekjaar As String

EurBefJaar = False
If Right(BOOKYEAR_FROMTO, 4) <> "1231" Then
    If bhEuro Then
        'eventjes kijken naar vorig jaar
        bClose TABLE_COUNTERS
        tempoBoekjaar = bstNaam(TABLE_COUNTERS)
        bstNaam(TABLE_COUNTERS) = "jr" & BJPERDAT.Boekjaar.Text - 1
        If String99(READING, 296) <> "EUR" Then
            'extra controle nodig...
            EurBefJaar = True
            MsgBox "Werkelijke jaar met documenten in EUR en BEF.", vbInformation
        End If
        bClose TABLE_COUNTERS
        bstNaam(TABLE_COUNTERS) = tempoBoekjaar
    End If
End If

End Function
