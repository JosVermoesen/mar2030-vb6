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
   Begin VB.CheckBox CheckBoxForceExecutionDate 
      Caption         =   "Forceer Memodatum"
      Enabled         =   0   'False
      Height          =   345
      Left            =   1680
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   360
      Width           =   2115
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
      Left            =   5340
      TabIndex        =   24
      Top             =   1320
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
      Left            =   9000
      TabIndex        =   23
      Top             =   1260
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
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
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
      Top             =   5310
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
      Left            =   10260
      Picture         =   "frmOGM.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   900
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
      Left            =   3780
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1125
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   3780
      TabIndex        =   3
      Top             =   960
      Width           =   1635
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   5460
      TabIndex        =   4
      Top             =   960
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
      Top             =   840
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
      Top             =   1080
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
      Left            =   9000
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   900
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
      Left            =   9000
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1620
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
      Format          =   16580611
      CurrentDate     =   36327
      MinDate         =   35796
   End
   Begin MSComCtl2.DTPicker DatumVerwerking 
      Height          =   315
      Left            =   120
      TabIndex        =   26
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
      CurrentDate     =   46023
      MaxDate         =   58862
      MinDate         =   46023
   End
   Begin MSComCtl2.DTPicker DTPickerGlobalMemoDate 
      Height          =   315
      Left            =   3840
      TabIndex        =   28
      Top             =   360
      Visible         =   0   'False
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
      CurrentDate     =   46023
      MaxDate         =   58862
      MinDate         =   46023
   End
   Begin VB.Label Label1 
      Caption         =   "Groep ID"
      Height          =   240
      Left            =   8760
      TabIndex        =   29
      Top             =   120
      Width           =   855
   End
   Begin VB.Label GuidLabel 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "GuidLabel"
      Height          =   255
      Left            =   8760
      TabIndex        =   25
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "EUR"
      Height          =   315
      Left            =   8220
      TabIndex        =   22
      Top             =   5340
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BEF"
      Height          =   315
      Left            =   8220
      TabIndex        =   21
      Top             =   5040
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblEUR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8760
      TabIndex        =   20
      Top             =   5340
      Width           =   1395
   End
   Begin VB.Label lblBEF 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8760
      TabIndex        =   19
      Top             =   5040
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lbBank 
      Caption         =   "Ban&k"
      Height          =   240
      Index           =   0
      Left            =   5400
      TabIndex        =   0
      Top             =   120
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
      Caption         =   "Datu&m"
      Height          =   240
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   735
   End
   Begin VB.Label lbVanTot 
      Caption         =   "&Van - Tot"
      Height          =   240
      Left            =   3840
      TabIndex        =   2
      Top             =   720
      Width           =   840
   End
End
Attribute VB_Name = "frmOGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim rsAny As ADODB.Recordset

Dim psTekst(5) As String
Dim TotaalD         As Long
Dim TotaalC         As Long

Dim VeldTXT(17) As String

Dim TLijnen             As Integer
Dim invalidAA   As Variant

'MAIN
Dim bancGuid As String 'ok

Dim guidId As String 'ok
Dim creationDateTime As String 'ok

Dim numberOfTransactions As Integer 'ok
Dim controlSum As String 'ok

Dim initiatingPartyName As String 'ok
Dim paymentInformationId As String
Dim requestedExecutionDate As String

Dim debtorName As String 'ok
Dim debtorIBAN As String 'ok
Dim debtorBIC As String 'ok
Dim transactionsList As String

Private Sub InitVelden()
Dim T As Integer

REPORT_FIELD(0) = "Lijn"
    REPORT_TAB(0) = 5

REPORT_FIELD(1) = "MemoDatum"
    REPORT_TAB(1) = 10

REPORT_FIELD(2) = "    Bedrag"
    REPORT_TAB(2) = 21

REPORT_FIELD(3) = "Munt"
    REPORT_TAB(3) = 32

REPORT_FIELD(4) = "Begunstigde"
    REPORT_TAB(4) = 37

REPORT_FIELD(5) = "Rekeningnr."
    REPORT_TAB(5) = 68

REPORT_FIELD(6) = "OGM/Referte"
    REPORT_TAB(6) = 87

REPORT_FIELD(7) = "DocumentID"
    REPORT_TAB(7) = 102

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
    Printer.Print Tab(5); psTekst(2); Tab(112); "Pagina : " + Dec$((PAGE_COUNTER), "##########");

    Printer.Print Tab(112); "Datum  : "; psTekst(0); vbCrLf; vbCrLf;
    Printer.Print Tab(5); UCase$(psTekst(3));

Printer.Print vbCrLf; Tab(5); FULL_LINE

Do While REPORT_TAB(T) <> 0
    Printer.Print Tab(REPORT_TAB(T));
    Printer.Print REPORT_FIELD(T);
    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
        Printer.Print vbCrLf;
    End If
    T = T + 1
Loop

Printer.Print Tab(5); FULL_LINE; vbCrLf$; vbCrLf$;

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
    Printer.Print vbCrLf; Tab(5); FULL_LINE
End If
For T = 0 To 6
    VeldTXT(T) = ""
Next

'VeldTXT(1) = "Totaal BEF"
'VeldTXT(2) = Dec(CDbl(lblBEF), MASK_BEF)
'T = 0
'aa = ""
'Do While T < 8
'    If chkAfdrukInVenster Then
'        aa = aa + VeldTXT(T) + vbTab
'    Else
'        Printer.Print ; Tab(REPORT_TAB(T));
'        Printer.Print ; VeldTXT(T);
'    End If
'    If REPORT_TAB(T + 1) < REPORT_TAB(T) Then
'        If chkAfdrukInVenster Then
'        Else
'            Printer.Print vbCrLf;
'        End If
'    End If
'    T = T + 1
'Loop
'If chkAfdrukInVenster Then Xlog.X.AddItem aa, Xlog.X.Rows - 1

VeldTXT(1) = "Totaal EUR"
VeldTXT(2) = Dec(CDbl(lblEUR), MASK_EUR)
VeldTXT(7) = GuidLabel.Caption
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
    
    guidId = Me.GuidLabel.Caption
    
    controlSum = Dec(lblEUR, ".00")
    numberOfTransactions = Trim(Str(grdDokumentDetail.Rows - 2))
    paymentInformationId = Mid(CreateGUID(), 2, 13)
    
    If Me.CheckBoxForceExecutionDate.Value = vbChecked Then
         requestedExecutionDate = Me.DTPickerGlobalMemoDate.Year & "-"
         requestedExecutionDate = requestedExecutionDate & Dec(Me.DTPickerGlobalMemoDate.Month, "00") & "-"
         requestedExecutionDate = requestedExecutionDate & Dec(Me.DTPickerGlobalMemoDate.Day, "00")
    Else
        MsgBox "TODO: check all executiondates"
        Exit Function
    End If
    
    Dim documentTemplate As String
    
    Ktrl = ScrLeesBestandAlleTekst(documentTemplate, PROGRAM_LOCATION + "xml-templates\sepa\be-sepa-main.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
    
    documentTemplate = Replace(documentTemplate, "{guidId}", guidId)
    documentTemplate = Replace(documentTemplate, "{creationDateTime}", creationDateTime)
    documentTemplate = Replace(documentTemplate, "{numberOfTransactions}", numberOfTransactions)
    documentTemplate = Replace(documentTemplate, "{controlSum}", controlSum)
    documentTemplate = Replace(documentTemplate, "{initiatingPartyName}", initiatingPartyName)
    documentTemplate = Replace(documentTemplate, "{paymentInformationId}", paymentInformationId)
    documentTemplate = Replace(documentTemplate, "{requestedExecutionDate}", requestedExecutionDate)
    documentTemplate = Replace(documentTemplate, "{debtorName}", debtorName)
    documentTemplate = Replace(documentTemplate, "{debtorIBAN}", debtorIBAN)
    documentTemplate = Replace(documentTemplate, "{debtorBIC}", debtorBIC)
        
    Dim templateNoOGM As String
    Dim templateNoRef As String
    Dim templateWithOGM As String
                    
    Ktrl = ScrLeesBestandAlleTekst(templateNoOGM, PROGRAM_LOCATION + "xml-templates\sepa\be-sepa-trans-no-ogm.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
    Ktrl = ScrLeesBestandAlleTekst(templateWithOGM, PROGRAM_LOCATION + "xml-templates\sepa\be-sepa-trans-ogm.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
    Ktrl = ScrLeesBestandAlleTekst(templateNoRef, PROGRAM_LOCATION + "xml-templates\sepa\be-sepa-trans-no-ref.xml")
    If Ktrl = 0 Then
        MsgBox "Onverwachte situatie", vbCritical
    End If
    
    Dim listOfTransActions As String
    Dim thisTransaction As String
    
    Dim linePayRef As String
    Dim lineSerialNumber As String
    Dim lineEndToEndId As String
    Dim lineAmount As String
    Dim lineCreditorName As String
    Dim lineCreditorIban As String
    Dim lineReference As String
       
    Dim volgNR As Integer
    For volgNR = 1 To grdDokumentDetail.Rows - 1
        If Len(grdDokumentDetail.TextMatrix(volgNR, 3)) <> 10 Then
            Exit For
        End If
        
        lineSerialNumber = Mid(GuidLabel, 9) + "+" + Format(volgNR, "0000")
        lineCreditorName = grdDokumentDetail.TextMatrix(volgNR, 1)
        lineEndToEndId = grdDokumentDetail.TextMatrix(volgNR, 2) 'Internal A document (Seller)
        lineAmount = Dec(grdDokumentDetail.TextMatrix(volgNR, 5), "")
        lineCreditorIban = grdDokumentDetail.TextMatrix(volgNR, 6)
        
        'TODO: check betaalreferte if it is OGM
        linePayRef = IbanCheck(grdDokumentDetail.TextMatrix(volgNR, 7), False, False)
        If linePayRef = grdDokumentDetail.TextMatrix(volgNR, 7) Then
            'OGM
            thisTransaction = templateWithOGM
            thisTransaction = Replace(thisTransaction, "{serialNumber}", lineSerialNumber)
            thisTransaction = Replace(thisTransaction, "{endToEndId}", lineEndToEndId)
            thisTransaction = Replace(thisTransaction, "{amount}", lineAmount)
            thisTransaction = Replace(thisTransaction, "{creditorName}", lineCreditorName)
            thisTransaction = Replace(thisTransaction, "{creditorIBAN}", lineCreditorIban)
            thisTransaction = Replace(thisTransaction, "{ogmReference}", linePayRef)
            
            listOfTransActions = listOfTransActions & thisTransaction & vbCrLf
        ElseIf Trim(grdDokumentDetail.TextMatrix(volgNR, 7)) = "" Then
            'NO REF!
            thisTransaction = templateNoRef
            thisTransaction = Replace(thisTransaction, "{serialNumber}", lineSerialNumber)
            thisTransaction = Replace(thisTransaction, "{endToEndId}", lineEndToEndId)
            thisTransaction = Replace(thisTransaction, "{amount}", lineAmount)
            thisTransaction = Replace(thisTransaction, "{creditorName}", lineCreditorName)
            thisTransaction = Replace(thisTransaction, "{creditorIBAN}", lineCreditorIban)
            
            listOfTransActions = listOfTransActions & thisTransaction & vbCrLf
        Else
            'NO OGM
            thisTransaction = templateNoOGM
            
            thisTransaction = Replace(thisTransaction, "{serialNumber}", lineSerialNumber)
            thisTransaction = Replace(thisTransaction, "{endToEndId}", lineEndToEndId)
            thisTransaction = Replace(thisTransaction, "{amount}", lineAmount)
            thisTransaction = Replace(thisTransaction, "{creditorName}", lineCreditorName)
            thisTransaction = Replace(thisTransaction, "{creditorIBAN}", lineCreditorIban)
            thisTransaction = Replace(thisTransaction, "{manualReference}", Trim(grdDokumentDetail.TextMatrix(volgNR, 7)))
            
            listOfTransActions = listOfTransActions & thisTransaction & vbCrLf
        End If
    Next
    documentTemplate = Replace(documentTemplate, "<Vsoft>transactionsList</Vsoft>", listOfTransActions)

    Dim ret As Integer
    Dim path As String * 260
    Dim desktopLocatie As String
    
    'Desktop lokatie
    ret = SHGetFolderPath(0, 0, 0, 0, path)
    desktopLocatie = Left(path, InStr(path, Chr(0)) - 1)
       
    Ktrl = ScrMaakTekstBestand(documentTemplate, LOCATION_COMPANYDATA + "coda\out\" + GuidLabel.Caption + ".xml")
    If Ktrl = 0 Then
        MsgBox "Fout bij bewaren van " + GuidLabel.Caption + ".xml", vbCritical
    Else
        Ktrl = ScrMaakTekstBestand(documentTemplate, desktopLocatie + "\" + GuidLabel.Caption + ".xml")
        If Ktrl = 0 Then
            MsgBox "Fout bij bewaren van " + GuidLabel.Caption + ".xml", vbCritical
        End If
        Msg = "Op uw bureaublad vindt U .xda bestand(en) klaar voor import in de toepassing van uw bank" & vbCrLf
        Msg = Msg & "VERGEET DEZE NIET te verwijderen na succesvol ondertekenen in de toepassing van uw bank" & vbCrLf & vbCrLf
        Msg = Msg & "Copijen blijven steeds behouden in uw bedrijfsinhoudsopgave"
        MsgBox Msg, vbExclamation + vbInformation
               
    End If
    
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
        debtorIBAN = Trim(vBibTekst(TABLE_VARIOUS, "#v259 #"))
        debtorBIC = Trim(vBibTekst(TABLE_VARIOUS, "#v260 #"))
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
        debtorIBAN = Trim(vBibTekst(TABLE_VARIOUS, "#v259 #"))
        debtorBIC = Trim(vBibTekst(TABLE_VARIOUS, "#v260 #"))
        Me.GuidLabel.Caption = Trim(vBibTekst(TABLE_VARIOUS, "#v237 #")) & "-" & bancGuid
    End If

End Sub


Private Sub CheckBoxForceExecutionDate_Click()

    If CheckBoxForceExecutionDate.Value = vbChecked Then
        Me.DTPickerGlobalMemoDate.Visible = True
        Me.DTPickerGlobalMemoDate.Value = Me.DatumVerwerking.Value + 1
    Else
        Me.DTPickerGlobalMemoDate.Visible = False
    End If
    
End Sub

Private Sub CmdEmailNBB_Click()

    Dim result As Boolean
    Dim volgNR As Integer

    Me.XmlOGM
    Drukken_Click
    MsgBox "STILL TESTING: Vlag en Guid toegevoegd voor latere opvolging.", vbExclamation
    
    For volgNR = 1 To grdDokumentDetail.Rows - 1
        If Len(grdDokumentDetail.TextMatrix(volgNR, 2)) = 0 Then
            Exit For
        End If
        result = ADO_GET(TABLE_INVOICES, 0, "=", grdDokumentDetail.TextMatrix(volgNR, 2))
        rsMAR(TABLE_INVOICES)("v411") = GuidLabel.Caption
        rsMAR(TABLE_INVOICES)("rvDM") = "0"
        rsMAR(TABLE_INVOICES)("dnnsync") = False
        rsMAR(TABLE_INVOICES).Update
    Next
    
    cmdSluiten_Click
        
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

    debtorName = Trim(String99(READING, 46))       'Naam opdrachtgever)
    debtorName = CheckforAmp(debtorName)
    initiatingPartyName = debtorName
    
    ' Adres niet meegeven is aanbevolen (BE) - Internationaal wel!
    ' String99(READING, 47)      'Adres opdrachtgever
    ' String99(READING, 48)       'Postnummer opdrachtgever
    ' Mid(String99(READING, 48), 6) 'Lokaliteit opdrachtgever
    
    Msg = "Deze module werkt voorlopig enkel met éénzelfde memodatum." & vbCrLf & vbCrLf
    Msg = Msg & "Datum + 1 of hoger (max. een toekomstige datum binnen het jaar!)"
    MsgBox Msg, vbExclamation, "XML betaalbestand aanmaken"
    
    
    dtpMemoDatum.MinDate = Now
    creationDateTime = GetCreationDateTime
    
    Me.DatumVerwerking.Value = MIM_GLOBAL_DATE
    CheckBoxForceExecutionDate.Value = vbChecked
    Me.DTPickerGlobalMemoDate.Value = Me.DatumVerwerking.Value + 1
    
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
            cbBank.AddItem vBibTekst(TABLE_VARIOUS, "#v231 #") + ": " + vBibTekst(TABLE_VARIOUS, "#v232 #") + " [" + vBibTekst(TABLE_VARIOUS, "#v259 #") + "]"
            bNext TABLE_VARIOUS
            If Ktrl Or Left(KEY_BUF(TABLE_VARIOUS), 2) <> "28" Then
                Exit Do
            End If
        Loop
        cbBank.ListIndex = 0
    End If
                
End Sub


Private Sub KTRLBalans(Fl As Integer)
    
    Dim VoorLetter  As String * 1
    Dim ktrlMemoDatum As String
    Dim GeenRekening As Integer
    Dim TeVerbeteren As Integer
    Dim aa          As Variant
    Dim dBetaald    As Double
    Dim dTotaal     As Double
    Dim checkerSEPA As String
        
    Set rsAny = New ADODB.Recordset

    invalidAA = ""
    If Fl = TABLE_SUPPLIERS Then
        VoorLetter = "L"
    Else
        VoorLetter = "K"
    End If

    ktrlMemoDatum = Format(Now + 1, "yyyymmdd")
    Screen.MousePointer = vbHourglass
   
    On Error Resume Next
    Err = 0
    rsAny.CursorLocation = adUseClient
    
    Msg = "SELECT Leveranciers.A110, Leveranciers.A100, Leveranciers.vs03, Leveranciers.v259, Leveranciers.e072, "
    Msg = Msg & "Dokumenten.v033, Dokumenten.v034, Dokumenten.v035, Dokumenten.v036, "
    Msg = Msg & "Dokumenten.v037, Dokumenten.v039, Dokumenten.v249, "
    Msg = Msg & "Dokumenten.v411, Dokumenten.rvDM, Dokumenten.rvID "
    Msg = Msg & "FROM Leveranciers, Dokumenten "
    Msg = Msg & "WHERE Left(Dokumenten.v033,2) = 'A0' "
    Msg = Msg & "AND Leveranciers.A110 >= '" & TekstLijn(2).text + "' "
    Msg = Msg & "AND Leveranciers.A110 <= '" & TekstLijn(3).text + "' "
    Msg = Msg & "AND Dokumenten.v034 = 'L' + Leveranciers.A110 "
    If Selektie(2).Value = 1 Then
        Msg = Msg & "AND Dokumenten.v035 >= '" & Left(BOOKYEAR_FROMTO, 8) & "' "
        Msg = Msg & "AND Dokumenten.v035 <= '" & Right(BOOKYEAR_FROMTO, 8) & "' "
    End If
    Msg = Msg & "ORDER BY Dokumenten.v034 "
    'Msg = Msg & "AND Val(Dokumenten.v037) <> Val(Dokumenten.v249) "
    SnelHelpPrint Msg, BL_LOGGING
    Screen.MousePointer = vbHourglass
    rsAny.Open Msg, adntDB, adOpenDynamic, adLockOptimistic
    Screen.MousePointer = vbNormal
    If Err Then
        MsgBox "Bron:" & vbCrLf & Err.Source & vbCrLf & vbCrLf & "Foutnummer: " & Err.Number & vbCrLf & vbCrLf & "Detail:" & vbCrLf & Err.Description
    ElseIf rsAny.RecordCount = 0 Then
        MsgBox "Geen documenten gevonden van/tot: " & TekstLijn(2) & " - " & TekstLijn(3)
        Exit Sub
    Else
        If grdDokumentDetail.Rows = 2 Then
        Else
            KtrlBox = MsgBox("Verrichtingen reeds aanwezig behouden." & vbCrLf & "Vermijdt dubbele bewerkingen!", vbQuestion + vbYesNo + vbDefaultButton2)
            If KtrlBox = vbNo Then
                GridSchoon grdDokumentDetail
            End If
        End If
        grdDokumentDetail.Refresh
        rsAny.MoveFirst
        Do While Not rsAny.EOF
            GoSub ValidateRecord
            rsAny.MoveNext
        Loop
        rsAny.Close
        Set rsAny = Nothing
    End If
    Screen.MousePointer = vbNormal
    
    If GeenRekening Then
        Msg = "Er zijn " & Str$(GeenRekening) & " verrichtingen zonder rekeningnummer" & vbCrLf
        Msg = Msg & invalidAA & vbCrLf
        MsgBox Msg, vbInformation
        Me.CmdEmailNBB.Enabled = True
        'KtrlBox = MsgBox(Msg, vbQuestion + vbDefaultButton2 + vbYesNo)
        'If KtrlBox = vbYes Then
        '    COUNT_TO = 1
        '    CmdEmailNBB.Enabled = True
        '    Do While COUNT_TO < grdDokumentDetail.Rows - 1
        '        'Debug.Print grdDokumentDetail.TextMatrix(COUNT_TO, 6)
        '        If Mid(grdDokumentDetail.TextMatrix(COUNT_TO, 6), 1, 2) = "!!" Then
        '            grdDokumentDetail.RemoveItem COUNT_TO
        '        Else
        '            COUNT_TO = COUNT_TO + 1
        '        End If
        '    Loop
        'End If
    Else
        Me.CmdEmailNBB.Enabled = True
    End If
    If TeVerbeteren Then
        MsgBox "Er zijn " & Str$(TeVerbeteren) & " rekeningnummers onjuist", vbInformation
    End If
    
    If aa = "" Then
        Beep
        MsgBox "Selectie buiten mogelijke documenten"
        Exit Sub
    End If

    'MsgBox "stop voor enkel leveranciers"
    Exit Sub

ValidateRecord:
    dBetaald = Val(objectValue(rsAny("v037")))
    dTotaal = Val(objectValue(rsAny("v249")))
    If dBetaald = dTotaal Then Return
    If objectValue(rsAny("rvDM")) = "0" Then
        Return
    End If
    If objectValue(rsAny("e072")) = "1" Then
        Return
    End If

    'TO DO creditnota's leveranciers, facturen klanten en kwitanties klanten uitsluiten !!
    Select Case Mid(rsAny("v033"), 1, 2)
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
        If rsAny("v035") >= Left(BOOKYEAR_FROMTO, 8) And rsAny("v035") <= Right(BOOKYEAR_FROMTO, 8) Then
        Else
            Return
        End If
    End If

    aa = rsAny("v034") & vbTab
    Select Case rsAny("vs03")
        Case "EUR"
        Case Else
            Return
    End Select
    aa = aa & objectValue(rsAny("A100")) & vbTab
    aa = aa & objectValue(rsAny("v033")) & vbTab
    'vervaldag < memodatum?
    If ktrlMemoDatum > objectValue(rsAny("v036")) Then
        aa = aa & DATE_TEXT(ktrlMemoDatum) & vbTab
    Else
        aa = aa & DATE_TEXT(objectValue(rsAny("v036"))) & vbTab
    End If
    aa = aa & objectValue(rsAny("vs03")) & vbTab
    If objectValue(rsAny("vs03")) = "EUR" And bhEuro = False Then
        aa = aa & Round((dTotaal - dBetaald) / EURO, 2) & vbTab
    ElseIf objectValue(rsAny("vs03")) = "BEF" And bhEuro = True Then
        aa = aa & Round((dTotaal - dBetaald) * EURO, 0) & vbTab
    ElseIf bhEuro = True Then
        'MsgBox "kontrolestop"
        If Round((dTotaal - dBetaald), 2) < 0 Then
            Return
        End If
        aa = aa & Round((dTotaal - dBetaald), 2) & vbTab
    ElseIf bhEuro = False Then
        aa = aa & Round((dTotaal - dBetaald), 0) & vbTab
    Else
        MsgBox "onlogische situatie"
    End If
    
    If Trim(objectValue(rsAny("v259"))) = "" Then
        aa = aa & "!! ..." & vbTab
        GeenRekening = GeenRekening + 1
    Else
        checkerSEPA = IbanCheck(objectValue(rsAny("v259")), True, False)
        If checkerSEPA = "invalid" Then
            aa = aa & "!!" + rsAny("v259") & vbTab
            TeVerbeteren = TeVerbeteren + 1
        Else
            aa = aa & rsAny("v259") & vbTab
        End If
    End If
    
    aa = aa & rsAny("v039") & vbTab
    On Local Error Resume Next
    aa = aa & objectValue(rsAny("rvID"))
    
    If Trim(objectValue(rsAny("v259"))) = "" Then
        invalidAA = invalidAA & aa & vbCrLf
    Else
        grdDokumentDetail.AddItem aa, grdDokumentDetail.Rows - 1
    End If
Return

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

    If Trim(TekstLijn(2)) = "0" Then
        bFirst TABLE_SUPPLIERS, 0
        TekstLijn(2).text = KEY_BUF(TABLE_SUPPLIERS)
        bLast TABLE_SUPPLIERS, 0
        TekstLijn(3).text = KEY_BUF(TABLE_SUPPLIERS)
    End If
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


