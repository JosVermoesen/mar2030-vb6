VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmOGM 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Betalingen/Overschrijvingen"
   ClientHeight    =   6060
   ClientLeft      =   960
   ClientTop       =   1320
   ClientWidth     =   11730
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6060
   ScaleWidth      =   11730
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Selektie 
      Caption         =   "&Verrichtingen in afwachting uitsluiten"
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
      Left            =   120
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1680
      Value           =   1  'Checked
      Width           =   3075
   End
   Begin VB.CheckBox cbLeveranciers 
      Alignment       =   1  'Right Justify
      Caption         =   "Enkel Leveranciers"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5940
      TabIndex        =   24
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1710
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Caption         =   "In venster"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8880
      TabIndex        =   23
      Top             =   1020
      Value           =   1  'Checked
      Width           =   1050
   End
   Begin VB.ComboBox cbBank 
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
      Left            =   4440
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3315
   End
   Begin MSFlexGridLib.MSFlexGrid grdDokumentDetail 
      Height          =   2835
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   5001
      _Version        =   393216
      FixedCols       =   0
      BackColor       =   -2147483624
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin VB.TextBox tbBedrag 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   4140
      TabIndex        =   11
      Top             =   5340
      Width           =   1515
   End
   Begin MSMask.MaskEdBox mebRekening 
      Height          =   315
      Left            =   60
      TabIndex        =   7
      Top             =   5340
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.CommandButton CmdEmailNBB 
      Caption         =   "&Versturen"
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
      Height          =   1035
      Left            =   10140
      Picture         =   "frmOGM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   660
      Width           =   1035
   End
   Begin VB.CommandButton Samenstellen 
      Caption         =   "Samenstellen"
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
      Left            =   4380
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1125
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   4380
      TabIndex        =   3
      Top             =   1080
      Width           =   1635
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   6060
      TabIndex        =   4
      Top             =   1080
      Width           =   1635
   End
   Begin VB.CheckBox Selektie 
      Caption         =   "C&ontrole Vervaldag"
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
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1200
      Value           =   1  'Checked
      Width           =   1755
   End
   Begin VB.CheckBox Selektie 
      Caption         =   "&Enkel dit boekjaar"
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
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1440
      Value           =   1  'Checked
      Width           =   1635
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&drukken"
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
      Left            =   8880
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   660
      Width           =   1125
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
      Height          =   300
      Left            =   8880
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1125
   End
   Begin MSComCtl2.DTPicker dtpMemoDatum 
      Height          =   315
      Left            =   2760
      TabIndex        =   9
      Top             =   5340
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
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
      Format          =   76546051
      CurrentDate     =   36327
      MinDate         =   35796
   End
   Begin MSComCtl2.DTPicker DatumVerwerking 
      Height          =   315
      Left            =   240
      TabIndex        =   27
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
      Format          =   76546051
      CurrentDate     =   46023
      MaxDate         =   58862
      MinDate         =   46023
   End
   Begin VB.Label GuidLabel 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GuidLabel"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EUR"
      Height          =   315
      Left            =   8220
      TabIndex        =   22
      Top             =   5040
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BEF"
      Height          =   315
      Left            =   8220
      TabIndex        =   21
      Top             =   5400
      Width           =   495
   End
   Begin VB.Label lblEUR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8760
      TabIndex        =   20
      Top             =   5040
      Width           =   1395
   End
   Begin VB.Label lblBEF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8760
      TabIndex        =   19
      Top             =   5400
      Width           =   1395
   End
   Begin VB.Label lbBank 
      Caption         =   "Ban&k"
      Height          =   240
      Index           =   0
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   600
   End
   Begin VB.Label lbBedrag 
      Alignment       =   1  'Right Justify
      Caption         =   "&Bedrag"
      Height          =   195
      Left            =   4800
      TabIndex        =   10
      Top             =   5100
      Width           =   795
   End
   Begin VB.Label lbReferte 
      Caption         =   "&Referte"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   5100
      Width           =   855
   End
   Begin VB.Label lbMemoDatum 
      Caption         =   "&Memodatum"
      Height          =   195
      Left            =   2820
      TabIndex        =   8
      Top             =   5100
      Width           =   1155
   End
   Begin VB.Label LabelGroupDate 
      Caption         =   "Groep ID en Datu&m"
      Height          =   240
      Left            =   240
      TabIndex        =   12
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lbVanTot 
      Caption         =   "&Van - Tot"
      Height          =   240
      Left            =   4440
      TabIndex        =   2
      Top             =   840
      Width           =   840
   End
End
Attribute VB_Name = "frmOGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim psTekst(5) As String
Dim TotaalD         As Long
Dim TotaalC         As Long

Dim VeldTXT(17) As String

Dim TLijnen             As Integer

Dim bancGuid As String


Private Sub InitVelden()
Dim T As Integer

REPORT_FIELD(0) = "Lijn"
    REPORT_TAB(0) = 2

REPORT_FIELD(1) = "MemoDatum"
    REPORT_TAB(1) = 7

REPORT_FIELD(2) = "    Bedrag"
    REPORT_TAB(2) = 18

REPORT_FIELD(3) = "Munt"
    REPORT_TAB(3) = 29

REPORT_FIELD(4) = "Begunstigde"
    REPORT_TAB(4) = 34

REPORT_FIELD(5) = "Rekeningnr."
    REPORT_TAB(5) = 65

REPORT_FIELD(6) = "OGM/Referte"
    REPORT_TAB(6) = 80

REPORT_FIELD(7) = "DocumentID"
    REPORT_TAB(7) = 95

REPORT_TAB(8) = 0

If chkAfdrukInVenster Then
    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = "Betalingen"
    Xlog.X.Cols = 7
    Xlog.X.Row = 0
    For T = 0 To 6
        Xlog.X.Col = T
        Xlog.X.text = REPORT_FIELD(T)
    Next
    Me.Show
End If

End Sub


Private Sub PrintTitel()
Dim T As Integer

'On Local Error GoTo PrtHandler3

If chkAfdrukInVenster Then Exit Sub
    
  If usrLicentieInfo <> "" Then
        Printer.CurrentX = 50
        Printer.CurrentY = 50
        Printer.Print usrLicentieInfo;
    End If
    PAGE_COUNTER = PAGE_COUNTER + 1
    Printer.CurrentY = 400
    Printer.Print Tab(1); psTekst(2); Tab(108); "Pagina : " + Dec$((PAGE_COUNTER), "##########");

    Printer.Print Tab(108); "Datum  : "; psTekst(0); vbCrLf; vbCrLf;
    Printer.Print Tab(1); UCase$(psTekst(3));

Printer.Print vbCrLf; FULL_LINE

Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print REPORT_FIELD(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop

Printer.Print FULL_LINE; vbCrLf$; vbCrLf$;

Exit Sub

PrtHandler3:
MsgBox "Kontroleer de printer."
Resume

End Sub


Private Sub PrintVelden()
Dim T As Integer
Dim aa As String

T = 0
aa = ""
Do While REPORT_TAB(T) <> 0
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print Tab(REPORT_TAB(T));
        Printer.Print VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If chkAfdrukInVenster Then
Else
    Printer.Print vbCrLf;
End If
If Printer.CurrentY >= Printer.Height - 1200 Then
    On Local Error Resume Next
    Printer.NewPage
    Printer.FontSize = Printer.FontSize
    Printer.Print " "
    Printer.FontSize = Printer.FontSize
    PrintTitel
End If
If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1

End Sub



Private Sub PrintTotaal()
Dim T As Integer
Dim aa As Variant

If grdDokumentDetail.Rows = 2 Then Exit Sub

If chkAfdrukInVenster Then
Else
    Printer.Print vbCrLf; FULL_LINE
End If
For T = 0 To 6
    VeldTXT(T) = ""
Next

VeldTXT(1) = "Totaal BEF"
VeldTXT(2) = Dec(CDbl(lblBEF), MASK_BEF)
T = 0
aa = ""
Do While T < 8
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print ; Tab(REPORT_TAB(T));
        Printer.Print ; VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1

VeldTXT(1) = "Totaal EUR"
VeldTXT(2) = Dec(CDbl(lblEUR), MASK_EUR)
T = 0
aa = ""
Do While T < 8
    If chkAfdrukInVenster Then
        aa = aa + VeldTXT(T) + vbTab
    Else
        Printer.Print ; Tab(REPORT_TAB(T));
        Printer.Print ; VeldTXT(T);
    End If
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        If chkAfdrukInVenster Then
        Else
            Printer.Print vbCrLf;
        End If
    End If
    T = T + 1
Loop
If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1

If chkAfdrukInVenster Then
    Xlog.X.Row = 1
    Xlog.X.Col = 0
    Xlog.X.ColWidth(0) = 555
    Xlog.X.ColWidth(1) = 1050
    Xlog.X.ColWidth(2) = 1020
    Xlog.X.ColWidth(3) = 495
    Xlog.X.ColWidth(4) = 2370
    Xlog.X.ColWidth(5) = 1275
    Xlog.X.ColWidth(6) = 1410
    Xlog.WindowState = 2
    
    Screen.MousePointer = vbNormal
    Xlog.WijzigenLijn.Visible = False
    Xlog.Afsluiten.Enabled = False
    Xlog.Afsluiten.TabStop = False
    Xlog.cbAfbeelding.Visible = False
    Xlog.SSTab1.TabVisible(1) = False
    XLogKey = ""
    Xlog.Show 1
    Unload Xlog
End If

End Sub


Function DosOGM()
   
    Dim boString As String * 128
    Dim goString1 As String * 128
    Dim goString2 As String * 128
    Dim eoString As String * 128
    Dim flHier As Integer
    
    Dim volgNr As Integer
    Dim bdrgTotaal As Currency
    Dim bdrgBedrag As Currency
    Dim rekNummer As Double
    Dim rekNummerTotaal As Double
    
    boString = "": goString1 = "": goString2 = "": eoString = ""
    Mim.Teken.InitDir = LOCATION_COMPANYDATA
    Mim.Teken.filename = ""
    Mim.Teken.ShowSave
    If Mim.Teken.filename = "" Then Exit Function
    flHier = FreeFile
    Open Mim.Teken.filename For Output As flHier

    Mid(boString, 1, 1) = "0"                       'beginopname
    Mid(boString, 4, 2) = "00"                      'voorwerp van betaling
    Mid(boString, 6, 6) = Format(Now, "ddmmyy")     'datum opmaak
    Mid(boString, 12, 3) = Left(cbBank.text, 3)     'codenummer fin.instelling
    Mid(boString, 15, 2) = "01"                     'toepassingscode 01= betaling 02=domicil 03=cheques
    Mid(boString, 17, 6) = Mid(boString, 6, 6)      'gevraagde uitvoeringsdatum
    Mid(boString, 24, 3) = "000"
    Mid(boString, 27, 12) = Mid(cbBank.text, InStr(cbBank.text, "[") + 1, 12)  'rekeningnummer opdrachtgever
    Mid(boString, 39, 26) = String99(READING, 46)      'Naam opdrachtgever
    Mid(boString, 65, 26) = String99(READING, 47)      'Adres opdrachtgever
    Mid(boString, 91, 4) = String99(READING, 48)       'Postnummer opdrachtgever
    Mid(boString, 95, 22) = Mid(String99(READING, 48), 6) 'Lokaliteit opdrachtgever
    Mid(boString, 117, 1) = "1"                     'Taalcode opdrachtgever
    Mid(boString, 128, 1) = "5"                     'Versiecode 5=EUR, 3=BEF
    Print #flHier, boString; vbNewLine;
    
    For volgNr = 1 To grdDokumentDetail.Rows - 1
        
        If Len(grdDokumentDetail.TextMatrix(volgNr, 3)) <> 10 Then
            Exit For
        End If
        
        'Gegevensopname 1
        Mid(goString1, 1, 1) = "1"                       'Identificatie opname
        Mid(goString1, 2, 4) = Format(volgNr, "0000")
        Mid(goString1, 6, 8) = Format(grdDokumentDetail.TextMatrix(volgNr, 8), "00000000") 'Refertenummer opdrachtgever
    
        rekNummer = Val(grdDokumentDetail.TextMatrix(volgNr, 6))
        rekNummerTotaal = rekNummerTotaal + rekNummer
        Mid(goString1, 24, 12) = grdDokumentDetail.TextMatrix(volgNr, 6) 'Rekeningnummer begunstigde
    
        'VeldTXT(1) = grdDokumentDetail.TextMatrix(Teller, 3)
    
        bdrgBedrag = grdDokumentDetail.TextMatrix(volgNr, 5)
        bdrgTotaal = bdrgTotaal + bdrgBedrag
        Mid(goString1, 36, 12) = Format(bdrgBedrag * 100, "000000000000") 'Bedrag
        Mid(goString1, 48, 26) = grdDokumentDetail.TextMatrix(volgNr, 1)
        Mid(goString1, 74, 1) = "1"                     'taalcode begunstigde
        Mid(goString1, 75, 12) = grdDokumentDetail.TextMatrix(volgNr, 7) 'Begin van de mededeling
        Mid(goString1, 87, 41) = ""                     'eerste vervolg gewone mededeling
        
        If Format(Val(grdDokumentDetail.TextMatrix(volgNr, 7)), "000000000000") <> grdDokumentDetail.TextMatrix(volgNr, 7) Then
            Mid(goString1, 128, 1) = "3" 'Aardcode 8= gestruct. mededeling anders 3
        Else
            If BankOk(Left(grdDokumentDetail.TextMatrix(volgNr, 7), 12)) Then
                Mid(goString1, 128, 1) = "8" 'Aardcode 8= gestruct. mededeling anders 3
            Else
                Mid(goString1, 128, 1) = "3" 'Aardcode 8= gestruct. mededeling anders 3
            End If
        End If
        Print #flHier, goString1; vbNewLine;
        
        'Gegevensopname 2
        Mid(goString2, 1, 1) = "2"                       'Identificatie opname
        Mid(goString2, 2, 4) = Format(volgNr, "0000")
        Mid(goString2, 6, 1) = "0"                       'code aanspreking begunstigde
        Mid(goString2, 7, 26) = ""                       'adres begunstigde
        Mid(goString2, 33, 4) = ""                       'postcode begunstigde
        Mid(goString2, 37, 22) = ""                      'Lokaliteit begunstigde
        Mid(goString2, 59, 53) = ""                      'mededeling: tweede vervolg
        Mid(goString2, 112, 1) = "0"                     'code chequekost 1= opdrgever, 2= begunstig.
        Print #flHier, goString2; vbNewLine;
        goString1 = "": goString2 = ""
        bGet TABLE_INVOICES, 0, grdDokumentDetail.TextMatrix(volgNr, 2)
            RecordToVeld TABLE_INVOICES
            vBib TABLE_INVOICES, "0", "rvDM"
            bUpdate TABLE_INVOICES, 0
    Next
    
    'Eindopname
    Mid(eoString, 1, 1) = "9"                       'Identificatie opname
    Mid(eoString, 2, 4) = Format((volgNr - 1) * 2, "0000")
    Mid(eoString, 6, 4) = Format((volgNr - 1), "0000")
    Mid(eoString, 10, 12) = Format(bdrgTotaal * 100, "000000000000")  'Totaal Bedragen
    Mid(eoString, 22, 15) = Format(rekNummerTotaal, "000000000000000")  'Totaal Rekeningnummers
    Mid(eoString, 37, 11) = "00" & Mid(String99(READING, 51), 1, 3) & Mid(String99(READING, 51), 5, 3) & Mid(String99(READING, 51), 9, 3)       'Identificatie afgever
    Print #flHier, eoString;
    Close flHier
    MsgBox Mim.Teken.filename & vbCrLf & vbCrLf & "staat klaar voor IMPORT door uw banksoftware.", vbInformation
    
End Function



Sub Herreken()

Dim TotaalBedraginBef As Currency
Dim Teller As Integer

TotaalBedraginBef = 0
For Teller = 1 To grdDokumentDetail.Rows - 1
    If grdDokumentDetail.TextMatrix(Teller, 4) = "EUR" Then
        TotaalBedraginBef = TotaalBedraginBef + (grdDokumentDetail.TextMatrix(Teller, 5) * EURO)
    ElseIf grdDokumentDetail.TextMatrix(Teller, 4) = "BEF" Then
        TotaalBedraginBef = TotaalBedraginBef + grdDokumentDetail.TextMatrix(Teller, 5)
    Else
        'Stop
    End If
Next
lblBEF = Format(Round(TotaalBedraginBef, 0), "#,###")
lblEUR = Format(Round(TotaalBedraginBef / EURO, 2), "#,###.00")
    
End Sub

Function XmlOGM()

    Dim bankSettingsKey As String
    bankSettingsKey = Trim("31" + Mid(cbBank.text, 1, 3))
    
    Stop
    vBib TABLE_VARIOUS, "31" + bankSettingsKey, "v005"
    vBib TABLE_VARIOUS, bancGuid, "v004"
    
    'vBib TABLE_VARIOUS, Left(KeuzeInfo(0).text, 4), "A010"
    '    vBib TABLE_VARIOUS, PolisNummer, "A000"
    '    vBib TABLE_VARIOUS, "K" + vBibTekst(TABLE_CUSTOMERS, "#A110 #"), "v004"
    
    '    If fKtrl = 99 Then
    '        bInsert TABLE_VARIOUS, 1
    '        If Ktrl Then MsgBox "stop"
    '    Else
    '        bUpdate TABLE_VARIOUS, 1
    '        If Ktrl Then MsgBox "stop"
    '    End If
    
End Function

Private Sub cbBank_Change()

    Dim bankSettingsKey As String
    bankSettingsKey = Trim("28" + Mid(cbBank.text, 1, 3))
    
    bGetOrGreater TABLE_VARIOUS, 1, bankSettingsKey
    
    If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 5) <> Trim(bankSettingsKey) Then
        MsgBox "Eerst financiële instelling parameters inbrengen via 'Diverse Gebruikersfiches' a.u.b.", vbInformation
        Unload Me
        Exit Sub
    Else
        RecordToVeld TABLE_VARIOUS
        bancGuid = Mid(CreateGUID(), 2, 13)
        Me.GuidLabel.Caption = Trim(vBibTekst(TABLE_VARIOUS, "#v237 #")) & "-" & bancGuid
    End If

End Sub

Private Sub cbBank_GotFocus()

    Dim bankSettingsKey As String
    bankSettingsKey = Trim("28" + Mid(cbBank.text, 1, 3))
    
    bGetOrGreater TABLE_VARIOUS, 1, bankSettingsKey
    
    If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 5) <> Trim(bankSettingsKey) Then
        MsgBox "Eerst financiële instelling parameters inbrengen via 'Diverse Gebruikersfiches' a.u.b.", vbInformation
        Unload Me
        Exit Sub
    Else
        RecordToVeld TABLE_VARIOUS
        bancGuid = Mid(CreateGUID(), 2, 13)
        Me.GuidLabel.Caption = Trim(vBibTekst(TABLE_VARIOUS, "#v237 #")) & "-" & bancGuid
    End If

End Sub


Private Sub CmdEmailNBB_Click()

    Me.XmlOGM
    'Me.DosOGM
    
End Sub

Private Sub cmdSluiten_Click()

    Unload Me

End Sub

Private Sub Drukken_Click()
Dim Teller As Integer

psTekst(2) = "Lijst Overschrijvingen " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
    psTekst(0) = MIM_GLOBAL_DATE
    psTekst(3) = "Bank " + cbBank

InitVelden
PAGE_COUNTER = 0
If chkAfdrukInVenster = 0 Then
    Set Printer = Printers(LijstPrinterNr)
    On Error Resume Next
    Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
    If Printer.Width > 12000 Then
        Printer.FontSize = 10
        Printer.FontName = "Courier New"
        Printer.Print " "
        Printer.FontSize = 10
    Else
        Printer.FontSize = 7.2
        Printer.FontName = "Courier New"
        Printer.Print " "
        Printer.FontSize = 7.2
    End If
End If
PrintTitel
For Teller = 1 To grdDokumentDetail.Rows - 1
    If Len(grdDokumentDetail.TextMatrix(Teller, 3)) <> 10 Then
        Exit For
    End If
    VeldTXT(0) = Format(Teller, "000")
    VeldTXT(1) = grdDokumentDetail.TextMatrix(Teller, 3)
    VeldTXT(3) = grdDokumentDetail.TextMatrix(Teller, 4)
    If VeldTXT(3) = "BEF" Then
        VeldTXT(2) = Dec((grdDokumentDetail.TextMatrix(Teller, 5)), MASK_BEF)
    Else
        VeldTXT(2) = Dec((grdDokumentDetail.TextMatrix(Teller, 5)), MASK_EUR)
    End If
    VeldTXT(4) = Left(grdDokumentDetail.TextMatrix(Teller, 1), 30)
    VeldTXT(5) = grdDokumentDetail.TextMatrix(Teller, 6)
    VeldTXT(6) = grdDokumentDetail.TextMatrix(Teller, 7)
    VeldTXT(7) = grdDokumentDetail.TextMatrix(Teller, 8)
    PrintVelden
Next
PrintTotaal
If chkAfdrukInVenster = 0 Then Printer.EndDoc
Screen.MousePointer = vbNormal
    
End Sub

Private Sub dtpMemoDatum_LostFocus()

grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 3) = Format(dtpMemoDatum, "dd/mm/yyyy")

End Sub

Private Sub Form_Load()
    
    MsgBox "Deze module wordt herwerkt. Niet gebruiken a.u.b.", vbExclamation
    
    dtpMemoDatum.MinDate = Now
    Me.DatumVerwerking.Value = MIM_GLOBAL_DATE

    TekstLijn(2).text = "0"
    TekstLijn(3).text = String$(12, "z")
    grdDokumentDetail.Rows = 2
    grdDokumentDetail.Cols = 9
    grdDokumentDetail.Row = 0
    
    grdDokumentDetail.Col = 0
    grdDokumentDetail.text = "ID Kode"
    grdDokumentDetail.ColWidth(0) = 1125
    
    grdDokumentDetail.Col = 1
    grdDokumentDetail.text = "Naam"
    grdDokumentDetail.ColWidth(1) = 1695
        
    grdDokumentDetail.Col = 2
    grdDokumentDetail.text = "Document"
    grdDokumentDetail.ColWidth(2) = 1155
        
    grdDokumentDetail.Col = 3
    grdDokumentDetail.text = "Memodatum"
    grdDokumentDetail.ColWidth(3) = 975
            
    grdDokumentDetail.Col = 4
    grdDokumentDetail.text = "Munt"
    grdDokumentDetail.ColWidth(4) = 450
    
    grdDokumentDetail.Col = 5
    grdDokumentDetail.text = "Bedrag"
    grdDokumentDetail.ColWidth(5) = 960
    
    grdDokumentDetail.Col = 6
    grdDokumentDetail.text = "SEPA Rekening"
    grdDokumentDetail.ColWidth(6) = 1620
        
    grdDokumentDetail.Col = 7
    grdDokumentDetail.text = "Referte/OGM"
    grdDokumentDetail.ColWidth(7) = 1380
        
    grdDokumentDetail.Col = 8
    grdDokumentDetail.text = "DocID"
    grdDokumentDetail.ColWidth(8) = 960
        
    bGetOrGreater TABLE_VARIOUS, 1, "28"
    If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "28" Then
        MsgBox "Eerst parameters van een financiële instelling inbrengen via 'Diverse Gebruikersfiches' a.u.b.", vbInformation
        Unload Me
        Exit Sub
    Else
        Do
            RecordToVeld TABLE_VARIOUS
            cbBank.AddItem vBibTekst(TABLE_VARIOUS, "#v231 #") + ": " + vBibTekst(TABLE_VARIOUS, "#v232 #") + " [" + vBibTekst(TABLE_VARIOUS, "#v236 #") + "]"
            bNext TABLE_VARIOUS
            If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "28" Then
                Exit Do
            End If
        Loop
        cbBank.ListIndex = 0
    End If
                
End Sub


Private Sub KTRLBalans(Fl As Integer)
Dim Cumul       As Double
Dim dTotaal     As Double
Dim dBetaald    As Double
Dim dBTW        As Double
Dim T           As Integer
Dim VoorLetter  As String * 1
Dim ktrlMemoDatum As String

Dim TeVerbeteren As Integer
Dim GeenRekening As Integer
Dim ZwareFout As Integer

Dim checkerSEPA As String

Dim AlgemeenTotaal As Currency

Dim aa          As Variant

If Fl = TABLE_SUPPLIERS Then
    VoorLetter = "L"
Else
    VoorLetter = "K"
End If

ktrlMemoDatum = Format(Now, "yyyymmdd")

T = 0
bGetOrGreater TABLE_INVOICES, 1, VoorLetter + TekstLijn(2)
If Ktrl Or KEY_BUF(TABLE_INVOICES) > VoorLetter + TekstLijn(3) Then
    Beep
    MsgBox "Selectie buiten mogelijke documenten"
    Exit Sub
End If
If VoorLetter + TekstLijn(2) = VoorLetter + TekstLijn(3) Then
    If KEY_BUF(TABLE_INVOICES) <> VoorLetter + Trim$(TekstLijn(2)) Then
        Beep
        MsgBox "Geen documenten voor " + vBibTekst(Fl, "#A100 #")
        Exit Sub
    End If
End If
    Screen.MousePointer = vbHourglass
    If grdDokumentDetail.Rows = 2 Then
    Else
        KtrlBox = MsgBox("Verrichtingen reeds aanwezig behouden.", vbQuestion + vbYesNo + vbDefaultButton2)
        If KtrlBox = vbNo Then
            GridSchoon grdDokumentDetail
        End If
    End If
    grdDokumentDetail.Refresh
    Do
        GoSub VolgendeLijn
        bNext TABLE_INVOICES
        If Ktrl Or KEY_BUF(TABLE_INVOICES) > "L" + TekstLijn(3) Then
            Exit Do
        End If
    Loop

Screen.MousePointer = vbNormal
If GeenRekening Then
    Me.CmdEmailNBB.Enabled = False
    Msg = "Er zijn " & Str$(GeenRekening) & " verrichtingen zonder rekeningnummer" & vbCrLf
    Msg = Msg & "Mogen deze verwijderd worden" & vbCrLf & vbCrLf
    Msg = Msg & "(zonder verwijdering enkel afdruk op papier mogelijk)"
    KtrlBox = MsgBox(Msg, vbQuestion + vbDefaultButton2 + vbYesNo)
    If KtrlBox = vbYes Then
        COUNT_TO = 1
        CmdEmailNBB.Enabled = True
        Do While COUNT_TO < grdDokumentDetail.Rows - 1
            'Debug.Print grdDokumentDetail.TextMatrix(COUNT_TO, 6)
            If Mid(grdDokumentDetail.TextMatrix(COUNT_TO, 6), 1, 2) = "!!" Then
                grdDokumentDetail.RemoveItem COUNT_TO
            Else
                COUNT_TO = COUNT_TO + 1
            End If
        Loop
    End If
Else
    Me.CmdEmailNBB.Enabled = True
End If
If TeVerbeteren Then
    MsgBox "Er zijn " & Str$(TeVerbeteren) & " rekeningnummers onjuist", vbInformation
End If
If ZwareFout Then
    MsgBox "Er zijn voor " & Str$(ZwareFout) & " identiteitsgegevens verdwenen", vbCritical, "Zware fout !!!"
End If
 
If aa = "" Then
    Beep
    MsgBox "Selectie buiten mogelijke documenten"
    Exit Sub
End If

'MsgBox "stop voor enkel leveranciers"
Exit Sub

VolgendeLijn:
RecordToVeld TABLE_INVOICES
dBetaald = Val(vBibTekst(TABLE_INVOICES, "#v037 #"))
dTotaal = Val(vBibTekst(TABLE_INVOICES, "#v249 #"))
If Round(dBetaald) = Round(dTotaal) Then Return
If vBibTekst(TABLE_INVOICES, "#rvDM #") = "0" And Me.Selektie(1).Value = vbChecked Then Return

'TO DO creditnota's leveranciers, facturen klanten en kwitanties klanten uitsluiten !!
Select Case Mid(vBibTekst(TABLE_INVOICES, "#v033 #"), 1, 2)
    Case "V0", "A1"  'Verkoopfactuur of Creditnota Aankoop
        'If Mid(vBibTekst(TABLE_INVOICES, "#v033 #"), 2, 1) = "1" Then
            'Return
        'End If
        Return
        
    Case "Q0" 'kwijting nog te betalen
        If dTotaal > 0 Then Return
    Case Else
        'Stop
End Select

'Selektie(0) = vervaldagcontrole
'TODO !

'Selektie(2)= enkel dit boekjaar
If Selektie(2).Value = 1 Then
    If rsMAR(TABLE_INVOICES)("v035") >= Left(BOOKYEAR_FROMTO, 8) And rsMAR(TABLE_INVOICES)("v035") <= Right(BOOKYEAR_FROMTO, 8) Then
    Else
        Return
    End If
End If

aa = vBibTekst(TABLE_INVOICES, "#v034 #") & vbTab
bGet Fl, 0, Mid(vBibTekst(TABLE_INVOICES, "#v034 #"), 2)
If Ktrl Then
    aa = aa & "Is verwijderd" & vbTab
    ZwareFout = ZwareFout + 1
Else
    RecordToVeld Fl
    Select Case vBibTekst(Fl, "#vs03 #")
        Case "EUR"
        Case Else
            Return
    End Select
    aa = aa & oWaarde(rsMAR(Fl)("A100")) & vbTab
End If
aa = aa & oWaarde(rsMAR(TABLE_INVOICES)("v033")) & vbTab
'vervaldag < memodatum?
If ktrlMemoDatum > oWaarde(rsMAR(TABLE_INVOICES)("v036")) Then
    aa = aa & DATE_TEXT(ktrlMemoDatum) & vbTab
Else
    aa = aa & DATE_TEXT(oWaarde(rsMAR(TABLE_INVOICES)("v036"))) & vbTab
End If
aa = aa & oWaarde(rsMAR(Fl)("vs03")) & vbTab
If oWaarde(rsMAR(Fl)("vs03")) = "EUR" And bhEuro = False Then
    aa = aa & Round((dTotaal - dBetaald) / EURO, 2) & vbTab
ElseIf oWaarde(rsMAR(Fl)("vs03")) = "BEF" And bhEuro = True Then
    aa = aa & Round((dTotaal - dBetaald) * EURO, 0) & vbTab
ElseIf bhEuro = True Then
    'MsgBox "kontrolestop"
    aa = aa & Round((dTotaal - dBetaald), 2) & vbTab
ElseIf bhEuro = False Then
    aa = aa & Round((dTotaal - dBetaald), 0) & vbTab
Else
    MsgBox "onlogische situatie"
End If

If Trim(oWaarde(rsMAR(Fl)("v259"))) = "" Then
    aa = aa & "!! ..." & vbTab
    GeenRekening = GeenRekening + 1
Else
    checkerSEPA = IbanCheck(oWaarde(rsMAR(Fl)("v259")), True, False)
    If checkerSEPA = "invalid" Then
        aa = aa & "!!" + vBibTekst(Fl, "#v259 #") & vbTab
        TeVerbeteren = TeVerbeteren + 1
    Else
        aa = aa & rsMAR(Fl)("v259") & vbTab
    End If
End If

aa = aa & vBibTekst(TABLE_INVOICES, "#v039 #") & vbTab
On Local Error Resume Next
aa = aa & oWaarde(rsMAR(TABLE_INVOICES)("rvID"))
grdDokumentDetail.AddItem aa, grdDokumentDetail.Rows - 1
Return

SluitAf:
Unload Xlog
Exit Sub

End Sub

Private Sub GrddokumentDetail_Click()

If grdDokumentDetail.Row = grdDokumentDetail.Rows - 1 Then
    mebRekening.Enabled = False
    dtpMemoDatum.Enabled = False
    tbBedrag.Enabled = False
Else
    bGet TABLE_SUPPLIERS, 0, Mid(grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 0), 2)
    If Ktrl Then
        MsgBox "Breng leverancier opnieuw in via fiches a.u.b.", vbCritical
    Else
        If Len(grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 6)) <> 16 Then
            MsgBox "Breng eerst 'juiste' rekeningnummer in a.u.b. voor " & vbCr & vbCr & grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 1), vbInformation
        End If
        RecordToVeld TABLE_SUPPLIERS
    
        If vBibTekst(TABLE_SUPPLIERS, "#v017 #") = "1" Then
            mebRekening.AutoTab = True
            mebRekening.ClipMode = mskExcludeLiterals
            mebRekening.PromptInclude = False
            mebRekening.Mask = "+++###/####/#####+++"
    
        Else
            mebRekening.AutoTab = False
            mebRekening.ClipMode = mskIncludeLiterals
            mebRekening.PromptInclude = True
            mebRekening.Mask = ""
        End If
        mebRekening.Enabled = True
        dtpMemoDatum.Enabled = True
        tbBedrag.Enabled = True
    End If
End If
mebRekening.text = grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 7)
tbBedrag.text = grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 5)
On Error Resume Next
dtpMemoDatum.Value = grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 3)

End Sub

Private Sub grddokumentDetail_GotFocus()

mebRekening.Enabled = False
dtpMemoDatum.Enabled = False
tbBedrag.Enabled = False

End Sub


Private Sub GrddokumentDetail_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 45 'Insert
        MsgBox KeyCode
        
    Case 46 'Delete
    If grdDokumentDetail.Row > 0 And grdDokumentDetail.Row < grdDokumentDetail.Rows - 1 Then
        grdDokumentDetail.RemoveItem grdDokumentDetail.Row
    End If
    grdDokumentDetail.SetFocus
    Herreken
    
End Select

End Sub

Private Sub grdDokumentDetail_RowColChange()

If grdDokumentDetail.Rows = 2 Then
Else
    GrddokumentDetail_Click
End If

End Sub

Private Sub mebRekening_LostFocus()

If mebRekening.Mask = "" Then
    grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 7) = mebRekening
ElseIf BankOk(mebRekening) Then
    grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 7) = mebRekening
Else
    mebRekening = grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 7)
End If

End Sub

Private Sub Samenstellen_Click()

    KTRLBalans TABLE_SUPPLIERS

    grdDokumentDetail.Row = 1
    grdDokumentDetail.Col = 0
    Herreken
    grdDokumentDetail.SetFocus

End Sub


Private Sub tbBedrag_LostFocus()

    If grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 5) = tbBedrag Then
    Else
        grdDokumentDetail.TextMatrix(grdDokumentDetail.Row, 5) = tbBedrag
        Herreken
    End If

End Sub

Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelLength = Len(TekstLijn(Index).text)
Select Case Index
    Case 0
        Samenstellen.Default = False
    Case 2, 3
        SnelHelpPrint "Dubbelklikken of [Ctrl] voor geïndexeerd zoeken", BL_LOGGING
        Samenstellen.Default = True
End Select

End Sub

Private Sub TekstLijn_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

Select Case Index
    Case 2, 3
        Select Case KeyCode
            Case 17
                aIndex = 0
                SharedFl = TABLE_SUPPLIERS
                GridText = TekstLijn(Index)
                SqlSearch.Show 1
                If Ktrl = 0 Then
                    TekstLijn(Index).text = vBibTekst(TABLE_SUPPLIERS, "#A110 #")
                    If Index = 2 Then
                        TekstLijn(3).text = TekstLijn(2).text
                    End If
                End If
        End Select
End Select

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)
Dim FlTemp As Integer

Select Case Index
    Case 1
        Samenstellen.Default = False
        If DATE_INVALID((TekstLijn(1).text)) Then
            Beep
            TekstLijn(1).text = MIM_GLOBAL_DATE
            TekstLijn(1).SetFocus
        End If
    Case 2, 3
        Samenstellen.Default = False
End Select

End Sub


