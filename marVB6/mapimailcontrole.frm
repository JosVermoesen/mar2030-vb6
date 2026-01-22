VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{0002E558-0000-0000-C000-000000000046}#1.0#0"; "OWC11.DLL"
Begin VB.Form frmMAPIMailCTRL 
   Caption         =   "R&Vsoft MailKontrole (lezen en verwijderen)"
   ClientHeight    =   7590
   ClientLeft      =   2925
   ClientTop       =   1590
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   10995
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   13150
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "mailBus"
      TabPicture(0)   =   "mapimailcontrole.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblAantalBerichten"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "RichTextBox1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdPostOphalen"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cbFetch"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdVerwijderen"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdSluiten"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmbBerichtType"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbBerichten"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "xml Data"
      TabPicture(1)   =   "mapimailcontrole.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Spreadsheet1"
      Tab(1).Control(1)=   "cbSendMail"
      Tab(1).Control(2)=   "cbIOMailKtrl"
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cbIOMailKtrl 
         Caption         =   "mailcontrole && bewaren"
         Height          =   375
         Left            =   -74880
         TabIndex        =   11
         Top             =   6960
         Width           =   2055
      End
      Begin VB.CommandButton cbSendMail 
         Caption         =   "Response Mail"
         Height          =   375
         Left            =   -72720
         TabIndex        =   10
         Top             =   6960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin OWC11.Spreadsheet Spreadsheet1 
         Height          =   6105
         Left            =   -74880
         OleObjectBlob   =   "mapimailcontrole.frx":0038
         TabIndex        =   9
         Top             =   360
         Width           =   10395
      End
      Begin VB.ComboBox lbBerichten 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   840
         Width           =   10455
      End
      Begin VB.ComboBox cmbBerichtType 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   480
         Width           =   8595
      End
      Begin VB.CommandButton cmdSluiten 
         Cancel          =   -1  'True
         Caption         =   "Sluiten"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdVerwijderen 
         Caption         =   "&Verwijderen"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cbFetch 
         Caption         =   "&Fetch"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   6960
         Width           =   1215
      End
      Begin VB.CommandButton cmdPostOphalen 
         Caption         =   "Post Ophalen"
         Height          =   375
         Left            =   1440
         TabIndex        =   1
         Top             =   6960
         Width           =   1215
      End
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   10515
         _ExtentX        =   18547
         _ExtentY        =   3836
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"mapimailcontrole.frx":083D
      End
      Begin VB.Label lblAantalBerichten 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   540
         Width           =   1215
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMBerichten 
      Left            =   3600
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISTaak 
      Left            =   2700
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "frmMAPIMailCTRL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cbFetch_Click()
    
    DoEvents
    MAPIMBerichten.Fetch
    DoEvents

End Sub

Private Sub cbIOMailKtrl_Click()

    Screen.MousePointer = vbHourglass
     
    Dim strSubject As String
    Dim strBody As String
    Dim Rij1 As String
    Dim Rij2 As String
    
    Dim strKop() As String
    Dim strVelden() As String
    Dim iPositie As Integer
    Dim strCell As String
    Dim TelTot2 As Integer
    Dim aPositie  As Integer
       
    Dim i As Integer
    Dim iSheetCount As Integer
    Dim SheetName As String

    iSheetCount = Spreadsheet1.ActiveWorkbook.Sheets.Count

    For i = 1 To iSheetCount
        SheetName = Spreadsheet1.ActiveWorkbook.Sheets(i).Name
        Spreadsheet1.Worksheets(SheetName).Activate
        'MsgBox SheetName
           
Nogeens:
        For TelTot = 0 To MAPIMBerichten.MsgCount - 1
            Err = 0
            On Local Error Resume Next
            MAPIMBerichten.MsgIndex = TelTot
            If Err Then MsgBox Error: Exit For
            strSubject = MAPIMBerichten.MsgSubject
            If strSubject = "$marWebIO$" & SheetName & "$marWebIO$" Then
                strBody = MAPIMBerichten.MsgNoteText
                iPositie = InStr(strBody, vbCrLf)
                If iPositie <> 0 Then
                    'ok
                    Rij1 = Mid(strBody, 1, iPositie - 1)
                    strKop = Split(Rij1, ";")
                    iPositie = iPositie + Len(vbCrLf)
                    strBody = Mid(strBody, iPositie)
                    Rij2 = Mid(strBody, 1, InStr(strBody, vbCrLf) - 1)
                    strVelden = Split(Rij2, ";")
                    If Spreadsheet1.ActiveSheet.Name = SheetName Then
                        'eerst de veldnamen invoegen
                        If Spreadsheet1.Range("A1").Value = Empty Then
                            For TelTot2 = 0 To UBound(strKop) - 1
                                If TelTot2 > 25 Then
                                    strCell = "A" & Chr(64 + TelTot2 - 26 + 1) & "1"
                                Else
                                    strCell = Chr(64 + TelTot2 + 1) & "1"
                                End If
                                Spreadsheet1.Range(strCell).Value = strKop(TelTot2)
                            Next
                        End If
                        'vervolgens positie achteraan bepalen
                        TelTot2 = 0
                        strCell = "A" & Trim(Str(TelTot2 + 1))
                        Do While Spreadsheet1.Range(strCell).Value <> Empty
                            TelTot2 = TelTot2 + 1
                            strCell = "A" & Trim(Str(TelTot2 + 1))
                        Loop
                        aPositie = TelTot2
                        'velden invoegen
                        For TelTot2 = 0 To UBound(strKop) - 1
                            If TelTot2 > 25 Then
                                strCell = "A" & Chr(64 + TelTot2 - 26 + 1) & Trim(Str(aPositie + 1))
                            Else
                                strCell = Chr(64 + TelTot2 + 1) & Trim(Str(aPositie + 1))
                            End If
                            Spreadsheet1.Range(strCell).Value = strVelden(TelTot2)
                        Next
                    Else
                        MsgBox SheetName & " sheet zoeken en actief plaatsen"
                    End If
                    'nu zou mailtje mogen verwijderd worden.
                    MAPIMBerichten.Delete mapMessageDelete
                    VernieuwBerichten
                    GoTo Nogeens
                End If
            End If
        Next
    Next
    Ktrl = ioBewaren
    Screen.MousePointer = vbNormal
       
End Sub

Private Sub cbSendMail_Click()

    
    Screen.MousePointer = vbHourglass
    
    Dim BestandHier As String
    Dim TelTot2 As Integer
    Dim strCell As String
    Dim Mailadres As String
    Dim colTrim As String
    Dim aantalRest As Integer
    Dim aantalMail As Integer
    
    If Me.MAPIMBerichten.SessionID = 0 Then Exit Sub
    On Local Error GoTo MPIError
    'Compose new message
    
    'If Spreadsheet1.ActiveSheet.Name = "Safety Jacket" Then
        'ok voor pilootfase
        'eerst de veldnamen invoegen
        TelTot2 = 0
        strCell = "A1"
        
        Do While Spreadsheet1.Range(strCell).Value <> Empty
            TelTot2 = TelTot2 + 1
            
            'controle voor mailadres
            colTrim = Trim(Str(TelTot2 + 1))
            strCell = "K" & colTrim 'kolom K bevat mailadres
            If Spreadsheet1.Range(strCell).Value = Empty Then
                aantalRest = aantalRest + 1
            Else
                Mailadres = Trim(Spreadsheet1.Range(strCell).Value)
                If Mailadres = "" Then
                    aantalRest = aantalRest + 1
                Else
                    aantalMail = aantalMail + 1
                    Me.MAPIMBerichten.Compose
                    'Address message
                    Me.MAPIMBerichten.RecipDisplayName = Me.Spreadsheet1.Range("A" & colTrim) & " " & Me.Spreadsheet1.Range("B" & colTrim)
                    'Me.MAPIMBerichten.RecipAddress = "SMTP:post@rv.be"
                    Me.MAPIMBerichten.RecipAddress = "SMTP:" & Mailadres
    
                    'Resolve recipient name
                    Me.MAPIMBerichten.AddressResolveUI = True '=dialogbox, false = error genereren
                    'MpiBericht.ResolveName
                    'Create the message
                    Me.MAPIMBerichten.MsgSubject = Spreadsheet1.ActiveSheet.Name
                    Msg = "Beste klanten," & vbCrLf & vbCrLf
                    Msg = Msg & "Uw gereserveerde fluovestje(s) zijn zojuist toegekomen.  "
                    Msg = Msg & "Ons kantoor is open maandag, dinsdag, donderdag en vrijdag 15.00 - 18.00 u., "
                    Msg = Msg & "behalve woensdag: 09.00 - 13.00 u.  Springt U eens binnen ?" & vbCrLf & vbCrLf
                    Msg = Msg & "Vestjes die niet afgehaald zijn voor 15/09 zullen wij automatisch tweede "
                    Msg = Msg & "helft deze maand aan huis bezorgen.  Dus, in ieder geval tot binnenkort !" & vbCrLf & vbCrLf
                    Msg = Msg & "Groetjes, Hedwig en Jos"
                
                    Me.MAPIMBerichten.MsgNoteText = Msg
                    'KtrlBox = MsgBox("Verstuurde E-mail aan " & vBibTekst(FlPartij, "#v224 #") & vbCr & vbCr & "Afdruk maken ?", vbQuestion + vbYesNo + vbDefaultButton2)
    
                    'Send the message
                    On Error Resume Next
                    Me.MAPIMBerichten.Send False
                    If Err <> 0 Then
                        Screen.MousePointer = vbNormal: Exit Sub
                    End If
                End If
            End If
            colTrim = Trim(Str(TelTot2 + 1))
            strCell = "A" & colTrim 'kolom A MOET iets bevatten
        Loop
        Msg = "mailverzendingen: " & aantalMail & vbCrLf & vbCrLf
        Msg = Msg & "andere: " & aantalRest
        MsgBox Msg
                  
Screen.MousePointer = vbNormal
Exit Sub

MPIError:
MsgBox Error
Resume Next

End Sub

Private Sub cmbBerichtType_Click()

VernieuwBerichten

End Sub

Private Sub cmdPostOphalen_Click()

MousePointer = vbHourglass
On Error Resume Next
MAPISTaak.SignOff
MAPISTaak.DownLoadMail = True
MAPISTaak.SignOn
MousePointer = vbNormal
If Err Then
    MsgBox Error
Else
    MAPIMBerichten.SessionID = MAPISTaak.SessionID
    VernieuwBerichten
    DoEvents
End If
Me.SetFocus

End Sub

Private Sub cmdSluiten_Click()

    Unload Me
    
End Sub

Private Sub cmdVerwijderen_Click()
    
jump:
    Msg = lbBerichten.Text & " verwijderen.  Bent U zeker" & vbCr & vbCr
    Msg = Msg & "Kies Ja om enkel dit bericht te verwijderen.  Nee voor ALLE BERICHTEN TE VERWIJDEREN !"
    KtrlBox = MsgBox(Msg, vbQuestion + vbYesNoCancel + vbDefaultButton3)
    If KtrlBox = vbYes Then
        MAPIMBerichten.Delete mapMessageDelete
        VernieuwBerichten
    ElseIf KtrlBox = vbNo Then
        Msg = "ALLE BERICHTEN verwijderen.  Bent U zeker" & vbCr & vbCr
        KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
        If KtrlBox = vbYes Then
            Do While MAPIMBerichten.MsgCount
                MAPIMBerichten.Delete mapMessageDelete
            Loop
            VernieuwBerichten
        End If
    Else
        GoTo jump
    End If
    
End Sub

Private Sub Form_Load()

Dim TelTot As Integer

On Error Resume Next
MAPISTaak.SignOn
If Err Then
    Unload Me
    Exit Sub
Else
    MAPIMBerichten.SessionID = MAPISTaak.SessionID
    cmbBerichtType.Clear
    cmbBerichtType.AddItem "IPM       : Normaal Bericht (Inbox)"
    cmbBerichtType.AddItem "IPC       : IPC Bericht (AS2)"
    cmbBerichtType.AddItem "IPC.AS2.E1: IPC Edifact Bericht (AS2)"
    cmbBerichtType.AddItem "IPC.AS2.P1: IPC positioneel Bericht (AS2)"
    cmbBerichtType.ListIndex = 0
End If
Ktrl = ioInladen
If MAPIMBerichten.MsgCount > 100 Then MsgBox "Uw Postvak IN bevat " & MAPIMBerichten.MsgCount & " berichten. Het is aanbevolen het aantal gelezen/ongelezen berichten in postvak IN zo laag mogelijk te houden.  TIP: Gebruik regels om belangrijke mail in submappen te plaatsen."

Me.SetFocus
Me.Refresh
 
End Sub


Private Sub Form_Resize()
    
    'Tabulatiepaneel aanpassen
    Me.SSTab1.Top = 0
    Me.SSTab1.Left = 0
    Me.SSTab1.Width = Me.Width - 100
    Me.SSTab1.Height = Me.Height - 550
    
    'SpreadSheet aanpassen
    Me.Spreadsheet1.Width = Me.Width - 300
    Me.Spreadsheet1.Height = Me.Height - 1800
    
    'Alle knoppen op juiste hoogte instellen
    Me.cbIOMailKtrl.Top = Me.Height - 1150
    Me.cbSendMail.Top = Me.Height - 1150
    Me.cbFetch.Top = Me.Height - 1150
    Me.cmdPostOphalen.Top = Me.Height - 1150
    Me.cmdSluiten.Top = Me.Height - 1150
    Me.cmdVerwijderen.Top = Me.Height - 1150
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error Resume Next
MAPISTaak.SignOff

End Sub

Private Sub lbBerichten_Click()

MAPIMBerichten.MsgIndex = lbBerichten.ListIndex
RichTextBox1.Text = MAPIMBerichten.MsgNoteText

End Sub

Function VernieuwBerichten()

MAPIMBerichten.FetchMsgType = Trim(Mid(cmbBerichtType.Text, 1, InStr(cmbBerichtType.Text, ":") - 1))
MAPIMBerichten.Fetch

lblAantalBerichten.Caption = MAPIMBerichten.MsgCount
lbBerichten.Clear
RichTextBox1.TextRTF = ""

If MAPIMBerichten.MsgCount > 0 Then
    cmdVerwijderen.Enabled = True
    
    For TelTot = 0 To MAPIMBerichten.MsgCount - 1
        Err = 0
        On Local Error Resume Next
        MAPIMBerichten.MsgIndex = TelTot
        If Err Then MsgBox Error: Exit For
        lbBerichten.AddItem Format(MAPIMBerichten.MsgDateReceived, "dd/mm/yyyy hh:mm ") & MAPIMBerichten.MsgSubject
    Next
    If lbBerichten.ListCount Then
        lbBerichten.ListIndex = 0
        On Error Resume Next
        lbBerichten.SetFocus
    End If
Else
    cmdVerwijderen.Enabled = False
End If

End Function

Function ioInladen()
    
    Dim TekstZelf As Variant

    KtrlBox = ScrLeesTekstBestand(TekstZelf, "ioWebMarNT.xml", BedrijfsLokatie, True)
    If KtrlBox = False Then
        MsgBox "Data kan niet ingeladen worden", vbExclamation
    Else
        'Spreadsheet1.XMLData = TekstZelf
    End If

End Function

Function ioBewaren()

    KtrlBox = ScrMaakTekstBestand(Spreadsheet1.XMLData, "ioWebMarNT.xml", BedrijfsLokatie)
    If KtrlBox = False Then
        MsgBox "Data kan niet bewaard worden", vbExclamation
    Else
        SnelHelpPrint "Data met succes weggeschreven"
    End If

End Function

Function ScrLeesTekstBestand(DataZelf As Variant, Bestandsnaam As String, Pathnaam As String, LeesAlles As Boolean) As Boolean
    
    On Error Resume Next
    Const ForReading = 1
    Dim A As TextStream
    
    Set A = fs.OpenTextFile(Pathnaam & "\" & Bestandsnaam, ForReading, False, TristateFalse)
    If Err Then ScrLeesTekstBestand = False: Exit Function
    
    If LeesAlles = True Then
        If UCase(Right(Bestandsnaam, 3)) = "XML" Then
            Spreadsheet1.XMLData = A.ReadAll
        Else
            Spreadsheet1.HTMLData = A.ReadAll
        End If
    Else
        Do While Not A.AtEndOfStream
            DataZelf = DataZelf & A.ReadLine & vbCr
        Loop
    End If
    A.Close
    If Err Then
        ScrLeesTekstBestand = False
    Else
        ScrLeesTekstBestand = True
    End If
    A = Nothing
    
End Function

Function ScrMaakTekstBestand(DataZelf As Variant, Bestandsnaam As String, Pathnaam As String) As Boolean

    On Error Resume Next
    Dim A As TextStream
    Screen.MousePointer = vbHourglass
    Set A = fs.CreateTextFile(Pathnaam & "\" & Bestandsnaam, True)
    A.WriteLine (DataZelf)
    A.Close
    If Err Then
        ScrMaakTekstBestand = False
    Else
        ScrMaakTekstBestand = True
    End If
    A = Nothing
    Screen.MousePointer = vbNormal
    
End Function


