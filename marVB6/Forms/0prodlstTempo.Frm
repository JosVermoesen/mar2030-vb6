VERSION 5.00
Begin VB.Form FrmLijstProdukten 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produktlijsten (rekenkundig)"
   ClientHeight    =   2430
   ClientLeft      =   180
   ClientTop       =   1485
   ClientWidth     =   6030
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
   ScaleHeight     =   2430
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox cbNulWaarden 
      Caption         =   "Stockwaarden cijfers <= 0 uitsluiten"
      Height          =   255
      Left            =   960
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1680
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CheckBox chkAfdrukLiggend 
      Alignment       =   1  'Right Justify
      Caption         =   "L&iggende Printerafdruk"
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2310
   End
   Begin VB.CheckBox chkAfdrukInVenster 
      Caption         =   "Afdruk in v&enster"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2040
      Value           =   1  'Checked
      Width           =   1830
   End
   Begin VB.ComboBox Sortering 
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
      ItemData        =   "0prodlstTempo.frx":0000
      Left            =   900
      List            =   "0prodlstTempo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   960
      Width           =   3555
   End
   Begin VB.TextBox TekstInfo 
      Height          =   315
      Index           =   1
      Left            =   2580
      TabIndex        =   8
      Text            =   "zzzzzzzzzzzzzzzzzzzz"
      Top             =   1290
      Width           =   1875
   End
   Begin VB.TextBox TekstInfo 
      Height          =   315
      Index           =   0
      Left            =   900
      TabIndex        =   7
      Top             =   1290
      Width           =   1635
   End
   Begin VB.ComboBox CmbLijstType 
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
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   630
      Width           =   3555
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   300
      Width           =   4320
   End
   Begin VB.TextBox TekstLijn 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   2880
      TabIndex        =   1
      Top             =   0
      Width           =   1560
   End
   Begin VB.CommandButton Drukken 
      Caption         =   "Af&drukken"
      Height          =   375
      Left            =   4530
      TabIndex        =   9
      Top             =   30
      Width           =   1440
   End
   Begin VB.CommandButton Annuleren 
      Cancel          =   -1  'True
      Caption         =   "&Sluiten"
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
      Left            =   4530
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   420
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Van - tot"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   1380
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sub&Titel"
      Height          =   195
      Index           =   0
      Left            =   270
      TabIndex        =   2
      Top             =   60
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Datu&m Drukken"
      Height          =   195
      Index           =   1
      Left            =   1470
      TabIndex        =   0
      Top             =   60
      Width           =   1335
   End
End
Attribute VB_Name = "FrmLijstProdukten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim psTekst(5) As String
Dim VeldTXT(17) As String
Dim aa As String

Dim TLijnen             As Long

Dim rsLijstProducten As ADODB.Recordset

Dim IndexNR As Integer
Dim tMaxVeld As Integer

Private Sub Annuleren_Click()
            
    Unload Me

End Sub


Private Sub chkAfdrukLiggend_Click()

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If chkAfdrukLiggend.Value = 1 Then
    Printer.Orientation = vbPRORLandscape
Else
    Printer.Orientation = vbPRORPortrait
End If

End Sub

Private Sub CmbLijstType_Click()

If CmbLijstType.ListIndex > 1 Then
    MsgBox "Voorlopig " & Now & " enkel eerste lijst mogelijk !"
    CmbLijstType.ListIndex = 0
End If

End Sub


Private Sub Drukken_Click()
Dim BeginSleutel    As String
Dim EindSleutel     As String
Dim Lijn            As Long
Dim DCBedrag        As Long
Dim XX As Integer
Dim Bedrag As Double

Dim TicketEUR2005 As Currency
Dim ticketEUR As Currency

Dim VeldTXT3 As String
Dim SQLstring As String

aa = ""
TLijnen = 0
Lijn = 0
BeginSleutel = TekstInfo(0)
EindSleutel = TekstInfo(1)

    psTekst(2) = CmbLijstType.text + " " + Mid(Mim.Caption, InStr(Mim.Caption, "["))
    psTekst(0) = TekstLijn(1).text
    psTekst(3) = TekstLijn(0).text

InitVelden
bClose TABLE_PRODUCTS
'Set rsLijstProducten = New ADODB.Recordset
'SQLstring = "SELECT * FROM Produkten WHERE "
'Stop

bFirst TABLE_PRODUCTS, IndexNR
bGetOrGreater TABLE_PRODUCTS, IndexNR, BeginSleutel
If Ktrl Or UCase(KEY_BUF(TABLE_PRODUCTS)) > UCase(EindSleutel) Then
    Beep
    Exit Sub
Else
    Screen.MousePointer = vbHourglass
    PAGE_COUNTER = 0
    If chkAfdrukInVenster = 0 Then
        If Printer.Width > 12000 Then
            Printer.FontSize = 10
            Printer.FontName = "Courier New"
            Printer.Print " "
            Printer.FontSize = 10
        Else
            Printer.FontSize = 8
            Printer.FontName = "Courier New"
            Printer.Print " "
            Printer.FontSize = 8
        End If
    End If
    PrintTitel
    GoSub PrintInfo
End If

Do
    bNext TABLE_PRODUCTS
    If Ktrl Or UCase(Trim$(KEY_BUF(TABLE_PRODUCTS))) > UCase(EindSleutel) Then
        Exit Do
    Else
        GoSub PrintInfo
    End If
Loop
PrintTotaal
If chkAfdrukInVenster = 0 Then
    Printer.EndDoc
End If
Screen.MousePointer = vbNormal
'Annuleren_Click
Exit Sub

Dim Basisbedrag As Currency
Dim BasisAantal As Single

PrintInfo:
RecordToVeld TABLE_PRODUCTS
Select Case CmbLijstType.ListIndex
    Case 0
        VeldTXT(0) = vBibTekst(TABLE_PRODUCTS, "#v102 #")
        VeldTXT(1) = vBibTekst(TABLE_PRODUCTS, "#v105 #")
        
        Basisbedrag = Val(vBibTekst(TABLE_PRODUCTS, "#e112 #"))
        VeldTXT(2) = Dec(Val(vBibTekst(TABLE_PRODUCTS, "#e112 #")), MASK_EUR & "00")
        
        VeldTXT(3) = vBibTekst(TABLE_PRODUCTS, "#v111 #")
        VeldTXT3 = Mid(fmarBoxText("002", "2", VeldTXT(3)), 4)
        
        BasisAantal = Val(vBibTekst(TABLE_PRODUCTS, "#v107 #"))
        VeldTXT(4) = Right(Dec(Val(vBibTekst(TABLE_PRODUCTS, "#v107 #")), MASK_SY(7)), 6)
        
        VeldTXT(5) = Mid(fmarBoxText("004", "2", vBibTekst(TABLE_PRODUCTS, "#v106 #")), 4)
        TicketEUR2005 = ((Basisbedrag * BasisAantal) + ((Basisbedrag * BasisAantal * Val(VeldTXT3) / 100)))
        VeldTXT(6) = Right(Dec((TicketEUR2005), MASK_EUR), 8)
        VeldTXT(7) = Right(Dec(Val(vBibTekst(TABLE_PRODUCTS, "#v114 #")) + Val(vBibTekst(TABLE_PRODUCTS, "#v119 #")) - Val(vBibTekst(TABLE_PRODUCTS, "#v120 #")), MASK_SY(7)), 6)
        VeldTXT(8) = vBibTekst(TABLE_PRODUCTS, "#v109 #")
        VeldTXT(9) = vBibTekst(TABLE_PRODUCTS, "#v125 #")
        
    Case 1
        VeldTXT(0) = vBibTekst(TABLE_PRODUCTS, "#v102 #")
        VeldTXT(1) = vBibTekst(TABLE_PRODUCTS, "#v105 #")
        
        Basisbedrag = Val(vBibTekst(TABLE_PRODUCTS, "#e113 #"))
        VeldTXT(2) = Dec(Val(vBibTekst(TABLE_PRODUCTS, "#e113 #")), MASK_EUR & "00")
        
        VeldTXT(3) = vBibTekst(TABLE_PRODUCTS, "#v111 #")
        VeldTXT3 = Mid(fmarBoxText("002", "2", VeldTXT(3)), 4)
        
        BasisAantal = Val(vBibTekst(TABLE_PRODUCTS, "#v107 #"))
        VeldTXT(4) = Right(Dec(Val(vBibTekst(TABLE_PRODUCTS, "#v107 #")), MASK_SY(7)), 6)
        VeldTXT(5) = Mid(fmarBoxText("004", "2", vBibTekst(TABLE_PRODUCTS, "#v106 #")), 4)
        
        'TicketEUR2005 = ((Basisbedrag * BasisAantal) + ((Basisbedrag * BasisAantal * Val(VeldTXT3) / 100)))
        'VeldTXT(6) = Right(Dec((TicketEUR2005), MASK_EUR), 8)
        VeldTXT(6) = Right(Dec(Val(vBibTekst(TABLE_PRODUCTS, "#v115 #")), MASK_SY(7)), 6)
        VeldTXT(7) = Right(Dec(Val(vBibTekst(TABLE_PRODUCTS, "#v114 #")) + Val(vBibTekst(TABLE_PRODUCTS, "#v119 #")) - Val(vBibTekst(TABLE_PRODUCTS, "#v120 #")), MASK_SY(7)), 6)
        VeldTXT(8) = Dec(Val(VeldTXT(6)) - Val(VeldTXT(7)), MASK_SY(7))
        VeldTXT(9) = vBibTekst(TABLE_PRODUCTS, "#v125 #")
End Select
PrintVelden
Return

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Dim T As Integer
Dim TempB_MODUS As Integer
Dim SorteringTel As Integer

TekstLijn(0).text = ""
TekstLijn(1).text = MIM_GLOBAL_DATE

Set Printer = Printers(LijstPrinterNr)
On Error Resume Next
Printer.PaperBin = LaadTekst(App.Title, "LIJSTPRINTER")
If Printer.Orientation = vbPRORLandscape Then
    chkAfdrukLiggend.Value = 1
Else
    chkAfdrukLiggend.Value = 0
End If
chkAfdrukLiggend_Click

CmbLijstType.AddItem "Lijst Verkoopstock"
CmbLijstType.AddItem "Lijst Te Bestellen"
CmbLijstType.AddItem "Lijst Aankoop stockwaarde"
CmbLijstType.ListIndex = 0

SorteringTel = -1
GetAllIndexes bstNaam(TABLE_PRODUCTS), Sortering
For T = 0 To Sortering.ListCount - 1
    If InStr(Sortering.List(T), FLINDEX_CAPTION(TABLE_PRODUCTS, 0)) Then
        Sortering.ListIndex = T
        Exit For
    End If
Next

End Sub

Private Sub InitVelden()
Dim T As Integer

Select Case CmbLijstType.ListIndex
    Case 0
        REPORT_FIELD(0) = "Nummer"
        REPORT_TAB(0) = 5

        REPORT_FIELD(1) = "Omschrijving"
        REPORT_TAB(1) = 19
        
        REPORT_FIELD(2) = "VK EUR Ex."
        REPORT_TAB(2) = 60
        
        REPORT_FIELD(3) = "B"
        REPORT_TAB(3) = 73
        
        REPORT_FIELD(4) = "Verpak"
        REPORT_TAB(4) = 75

        REPORT_FIELD(5) = "Maat"
        REPORT_TAB(5) = 83
    
        REPORT_FIELD(6) = "EUR incl"
        REPORT_TAB(6) = 89
        
        REPORT_FIELD(7) = " Stock"
        REPORT_TAB(7) = 98
        
        REPORT_FIELD(8) = "Plaats"
        REPORT_TAB(8) = 107
        
        REPORT_FIELD(9) = "Vlag"
        REPORT_TAB(9) = 120
        tMaxVeld = 9
        
    Case 1
        REPORT_FIELD(0) = "Nummer"
        REPORT_TAB(0) = 5

        REPORT_FIELD(1) = "Omschrijving"
        REPORT_TAB(1) = 19
        
        REPORT_FIELD(2) = "AK EUR Ex."
        REPORT_TAB(2) = 60
        
        REPORT_FIELD(3) = "B"
        REPORT_TAB(3) = 73
        
        REPORT_FIELD(4) = "Verpak"
        REPORT_TAB(4) = 75

        REPORT_FIELD(5) = "Maat"
        REPORT_TAB(5) = 83
    
        REPORT_FIELD(6) = "Min.Stock"
        REPORT_TAB(6) = 89
        
        REPORT_FIELD(7) = " Stock"
        REPORT_TAB(7) = 98
        
        REPORT_FIELD(8) = "Bestellen"
        REPORT_TAB(8) = 107
                
        REPORT_FIELD(9) = "Vlag"
        REPORT_TAB(9) = 120
        tMaxVeld = 9
        
    Case Else
        MsgBox "Stop"
End Select

If chkAfdrukInVenster Then
    Me.Hide
    Unload Xlog
    Xlog.Hide
    Xlog.Caption = CmbLijstType.text
    Xlog.X.Cols = tMaxVeld + 1
    Xlog.X.Row = 0
    For T = 0 To tMaxVeld
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
    Printer.Print Tab(1); psTekst(2); Tab(107); Dec((PAGE_COUNTER), "##########");

    Printer.Print Tab(107); psTekst(0); vbCrLf; vbCrLf;
    Printer.Print Tab(1); UCase$(psTekst(3));

Printer.Print vbCrLf$; FULL_LINE

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

Private Sub PrintTotaal()
Dim T As Integer

For T = 0 To tMaxVeld
    VeldTXT(T) = ""
Next
VeldTXT(1) = "Totaal aantal lijnen : " + Format(TLijnen)

If chkAfdrukInVenster Then
Else
    Printer.Print vbCrLf; FULL_LINE
End If

T = 0
aa = ""
Do While T < tMaxVeld
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
    Select Case CmbLijstType.ListIndex
        Case 0
        Xlog.X.Row = 1
        Xlog.X.Col = 0
        Xlog.X.ColWidth(0) = 1335
        Xlog.X.ColWidth(1) = 3360
        Xlog.X.ColWidth(2) = 810
        Xlog.X.ColWidth(3) = 210
        Xlog.X.ColWidth(4) = 615
        Xlog.X.ColWidth(5) = 525
        Xlog.X.ColWidth(6) = 945
        Xlog.X.ColWidth(7) = 900
        Xlog.X.ColWidth(8) = 1425
        Xlog.WindowState = 2

        Xlog.X.ColAlignment(2) = flexAlignRightTop
        Xlog.X.ColAlignment(4) = flexAlignRightTop
        Xlog.X.ColAlignment(6) = flexAlignRightTop
        Xlog.X.ColAlignment(7) = flexAlignRightTop
        Xlog.X.ColAlignment(8) = flexAlignRightTop
    End Select
        
    Screen.MousePointer = vbNormal
    Xlog.WijzigenLijn.Visible = False
    Xlog.Afsluiten.TabStop = False
    Xlog.Afsluiten.Enabled = False
    Xlog.cbAfbeelding.Visible = False
    XLogKey = ""
    Xlog.SSTab1.TabVisible(1) = False
    Xlog.Show 1
    'Unload Xlog
End If

End Sub

Private Sub PrintVelden()
Dim T As Integer
Dim aa As String

T = 0
aa = ""
Dim bstAantal As Single

If CmbLijstType.ListIndex = 1 Then
    'Bestellijst geselecteerd.  Enkel afdrukken indien nodig
    If Val(VeldTXT(9)) > 0 Then
    Else
        Exit Sub
    End If
End If
TLijnen = TLijnen + 1

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



Private Sub SorTering_Click()
Dim T As Integer

IndexNR = -1
For T = 0 To Sortering.ListCount - 1
    If InStr(Sortering.text, FLINDEX_CAPTION(TABLE_PRODUCTS, T)) Then
        IndexNR = T
        Exit For
    End If
Next
If IndexNR < 0 Then
    MsgBox "Indexen zijn ondertussen vernieuwd ?" + vbCrLf + vbCrLf + "Om van de nieuwe indexen gebruik te kunnen maken dient U het bedrijf te heropenen !"
End If

End Sub


Private Sub TekstLijn_GotFocus(Index As Integer)

TekstLijn(Index).SelLength = Len(TekstLijn(Index).text)

End Sub

Private Sub TekstLijn_LostFocus(Index As Integer)

Select Case Index
    Case 1
        If DATE_INVALID((TekstLijn(1).text)) Then
            Beep
            TekstLijn(1).text = MIM_GLOBAL_DATE
        End If
End Select

End Sub

