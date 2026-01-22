VERSION 5.00
Begin VB.Form LayOutpdfDokument 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VPE-PDF AfdrukTester"
   ClientHeight    =   3585
   ClientLeft      =   210
   ClientTop       =   645
   ClientWidth     =   5910
   ForeColor       =   &H80000008&
   HelpContextID   =   23000
   Icon            =   "0PDFLODOK.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   239
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   394
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtDemoTekst 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "Demotekst"
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton cbFont 
      Caption         =   "Vervolgens Font, Grootte en parameters"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   720
      Width           =   3615
   End
   Begin VB.TextBox txtFont 
      Height          =   405
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   3
      ToolTipText     =   "Kies kleur, font en parameters en daarna selecteren en copiëren voor uw eigen tekstlijnen"
      Top             =   1200
      Width           =   3735
   End
   Begin VB.CommandButton cbKleurKiezen 
      Caption         =   "Eerst Kleur Kiezen !"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox txtKleur 
      Height          =   405
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Teksthelper"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3000
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "DemoTekst"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   5535
   End
   Begin VB.Menu Bestanden 
      Caption         =   "&Bestand"
      Begin VB.Menu Bestand 
         Caption         =   "&Openen"
         Index           =   1
      End
      Begin VB.Menu Bestand 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Bestand 
         Caption         =   "Via kladblok"
         Index           =   3
         Begin VB.Menu ViaKB 
            Caption         =   "Tekst- en Lijnobjecten"
            Index           =   0
         End
      End
      Begin VB.Menu Bestand 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu Bestand 
         Caption         =   "&Afsluiten"
         Index           =   5
      End
   End
   Begin VB.Menu TaalKeuze 
      Caption         =   "&Taal"
      Begin VB.Menu mTaal 
         Caption         =   "&Frans"
         Index           =   1
      End
      Begin VB.Menu mTaal 
         Caption         =   "&Nederlands"
         Checked         =   -1  'True
         Index           =   2
      End
      Begin VB.Menu mTaal 
         Caption         =   "&Engels"
         Index           =   3
      End
      Begin VB.Menu mTaal 
         Caption         =   "&Duits"
         Index           =   4
      End
   End
   Begin VB.Menu Document 
      Caption         =   "Document"
      Begin VB.Menu DokType 
         Caption         =   "Factuur/Creditnota"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu DokType 
         Caption         =   "LeveringsBon"
         Index           =   1
      End
      Begin VB.Menu DokType 
         Caption         =   "BestelBon"
         Index           =   2
      End
      Begin VB.Menu DokType 
         Caption         =   "Offerte"
         Index           =   3
      End
      Begin VB.Menu DokType 
         Caption         =   "Briefwisseling"
         Index           =   4
      End
      Begin VB.Menu DokType 
         Caption         =   "Rekeninguitttreksel"
         Index           =   5
      End
      Begin VB.Menu DokType 
         Caption         =   "Kwijting"
         Index           =   6
      End
   End
End
Attribute VB_Name = "LayOutpdfDokument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim KopVoet As Integer
Dim TaalKode As Integer
Dim DocuType As Integer
Dim Diepgang As Integer

Dim psTekst(50) As String              'TekstString zelf
Dim psX(50) As Single                 'Xpositie
Dim psY(50) As Single                 'Ypositie
Dim psFontSize(50) As Single
Dim psFontName(50) As String
Dim psFontBold(50)         As Integer
Dim psFontItalic(50) As Integer
Dim psFontUnderLine(50) As Integer
Dim psColor(50) As Long
Dim MaxPslokatie As Integer                 'aantal STANDAARD af te drukken

Dim BoxTeller As Integer
Dim Box(10) As Integer
Dim BeginX(10) As Integer
Dim BeginY(10) As Integer
Dim TotX(10) As Integer
Dim TotY(10) As Integer

Dim PsLokatie As Integer

Dim FiguurTeller As Integer
Dim FiguurX(4) As Integer
Dim FiguurY(4) As Integer
Dim FiguurName(4) As String

Private Sub Bestand_Click(Index As Integer)

Select Case Index
    Case 1
        Inladen Format(KopVoet) + Format(TaalKode) + Format(DocuType)
        PsLokatie = 0
        
    Case 5
        Unload Me
    Case Else
End Select

End Sub

Private Sub cbFont_Click()

' Set Cancel to True
  Mim.Teken.CancelError = True
  On Error GoTo FontErrHandler
  ' Set the Flags property
  Mim.Teken.FLAGS = cdlCFEffects Or cdlCFBoth
    Mim.Teken.FontName = Label1.Font.Name
    Mim.Teken.FontSize = Label1.Font.Size
    Mim.Teken.FontBold = Label1.Font.Bold
    Mim.Teken.FontItalic = Label1.Font.Italic
  ' Display the Font dialog box
  Mim.Teken.ShowFont
  Label1.Font.Name = Mim.Teken.FontName
  Label1.Font.Size = Mim.Teken.FontSize
  Label1.Font.Bold = Mim.Teken.FontBold
  Label1.Font.Italic = Mim.Teken.FontItalic
  Me.txtFont.Text = Mim.Teken.FontSize & "," & """" & Mim.Teken.FontName & """" & "," & Me.Label1.ForeColor & ","
  If Mim.Teken.FontBold = True Then
    Me.txtFont.Text = Me.txtFont.Text & "1,"
  Else
    Me.txtFont.Text = Me.txtFont.Text & "0,"
  End If
  If Mim.Teken.FontItalic = True Then
    Me.txtFont.Text = Me.txtFont.Text & "1,"
  Else
    Me.txtFont.Text = Me.txtFont.Text & "0,"
  End If
  If Mim.Teken.FontUnderline = True Then
    Me.txtFont.Text = Me.txtFont.Text & "1"
  Else
    Me.txtFont.Text = Me.txtFont.Text & "0"
  End If
  'Label1.Font.Underline = Mim.Teken.FontUnderline
  'Label1.FontStrikethru = Mim.Teken.FontStrikethru
  'Label1.ForeColor = Mim.Teken.Color
  Exit Sub
FontErrHandler:
  ' User pressed the Cancel button
  Exit Sub

End Sub

Private Sub cbKleurKiezen_Click()

' Set Cancel to True
Mim.Teken.CancelError = True
  On Error GoTo ColorErrHandler
  'Set the Flags property
Mim.Teken.FLAGS = cdlCCRGBInit
  ' Display the Color Dialog box
Mim.Teken.ShowColor
  ' Set the form's background color to selected color
Me.Label1.ForeColor = Mim.Teken.Color
Me.txtKleur.Text = Mim.Teken.Color

ColorErrHandler:

End Sub

Private Sub DokType_Click(Index As Integer)

For T = 0 To 5
    DokType(T).Checked = False
Next

DokType(Index).Checked = True
DocuType = Index

End Sub

Private Sub Form_Load()

'If Not Toegankelijk(Me) Then
'    Unload Me
'    Exit Sub
'End If
KopVoet = 1
TaalKode = 2
Mim.WindowState = vbMinimized

Dim Bestanden As String
Me.txtKleur.Text = Me.Label1.BackColor

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

   If Mim.Report.IsOpen = True Then
        MsgBox "Sluit eerst het PDF venster a.u.b.", vbInformation
        Cancel = True
   End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

Close
Mim.WindowState = 0

End Sub

Private Sub Inladen(TypeEnTaal As String)

    Dim FlFree As Integer
    Dim Teller As Integer
    Dim pdfCmd As String

    Dim Ktrl As Double
    
    Dim vsoftVAN As Double
    Dim VsoftTot As Double
    
    Dim adresYpos As Double
    Dim adresXpos As Double
    Dim adresYpos2 As Double
    Dim adresXpos2 As Double
    Dim adresBox As Boolean
    
    Dim YPOS As Double
    Dim xPos As Double
    Dim YPOS2 As Double
    Dim xPos2 As Double
    Dim pdfFontSize As Double
    Dim pdfFontName As String
    Dim pdfFontBold As Double
    Dim pdfFontItalic As Double
    Dim pdfFontUnderLine As Double
    Dim pdfColor As Double
    Dim pdfAlign As Double
    Dim pdfBoxPen As Double
        
    Dim textstring As String
    Dim texttmp As String
      
    Dim COUNT_TO As Integer
        
    Mim.Report.CloseDoc
    
    
                            
VPETEST:
    With Mim.Report
        .OpenDoc
        .Author = Trim(String99(READING, 46))
        .GUILanguage = 3 'Nederlands
        .Title = "marIntegraal Rapport"
        '.nBottomMargin = 1 'deze margin geeft problemen
        .nTopMargin = 1
        .nLeftMargin = 1
        .nRightMargin = 1
        '.nBottomMargin = 1
        .GridVisible = True
    End With
    
    '.Top = 6
    '.Left = 165
    '.Width = 500
    '.Height = 560
    
If Dir(LOCATION_COMPANYDATA + "vpeSjbs\pdfDDEF" + TypeEnTaal + ".Txt") = "" Then
    Beep
    Mim.Report.CloseDoc
    MsgBox "vpeSjbs\pdfDDEF" + TypeEnTaal + ".Txt niet gevonden in de bedrijfsinhoudsopgave. Hierna wordt een voorbeelddocument vanuit de programmainhoudsopgave gecopiëerd", vbInformation
    
    If Not CopyFile(App.path & "\VpeSjbs", LOCATION_COMPANYDATA + "vpeSjbs\", "pdfDDEF" + TypeEnTaal + ".Txt") Then
        MsgBox App.path & "\VpeSjbs\pdfDDEF" + TypeEnTaal + ".Txt kan als voorbeelddocument niet vanuit de programmainhoudsopgave gecopiëerd worden.  Probeer eventueel manueel", vbInformation
    Else
        If Not CopyFile(App.path & "\VpeSjbs", LOCATION_COMPANYDATA, "demo-compagny-logo.bmp") Then
            MsgBox App.path & "\VpeSjbs\demo-compagny-logo" + ".bmp kan als voorbeeldafbeelding niet vanuit de programmainhoudsopgave gecopiëerd worden.  Probeer eventueel manueel", vbInformation
        End If
        GoTo VPETEST
    End If
    Exit Sub
Else
    FlFree = FreeFile
    Open LOCATION_COMPANYDATA + "vpeSjbs\pdfDDEF" + TypeEnTaal + ".Txt" For Input As FlFree
    
    Do While Not EOF(FlFree)
        Line Input #FlFree, pdfCmd
        If Left(pdfCmd, 1) = "'" Then
        Else
            Select Case Trim(UCase(pdfCmd))
                Case "CMD-VSOFTSPACE"
                    GoSub CMDVSOFTSPACE
                
                Case "CMD-ADRESSPACE"
                    GoSub CMDADRESSPACE
                
                Case "CMD-WRITE"
                    GoSub CMDWRITE
                
                Case "CMD-WRITEBOX"
                    GoSub CMDWRITEBOX
                
                Case "CMD-PRINT"
                    GoSub CMDPRINT
            
                Case "CMD-PICTURE"
                    GoSub CMDPICTURE
            
                Case Else
                    MsgBox pdfCmd & " nog niet voorzien", vbCritical
            End Select
        End If
    Loop
    Close FlFree
    'test voor barcode
    'Mim.Report.Barcode 1, 1, 3, 3, 1, "123456789012", ""
    Mim.Report.WriteDoc (PROGRAM_LOCATION + "marrapport.pdf")
    Mim.Report.Preview
    'Report.CloseDoc
End If

Exit Sub

CMDPICTURE:
Input #FlFree, xPos, YPOS, xPos2, YPOS2, pdfBoxPen, filename$
Mim.Report.PictureBestFit = True
'2.5, 0.5, 19, 2.5
If Left(filename$, 4) = "[BL]" Then
    filename$ = LOCATION_COMPANYDATA & Mid(filename$, 5)
ElseIf Left(filename$, 4) = "[PL]" Then
    filename$ = PROGRAM_LOCATION & Mid(filename$, 5)
End If
Ktrl = Mim.Report.Picture(xPos, YPOS, xPos2, YPOS2, filename$)
'.VpePrint(xPos, YPOS, textstring)
Return

CMDVSOFTSPACE:
Input #FlFree, vsoftVAN, VsoftTot
Mim.Report.FontName = "Courier New"
Mim.Report.FontSize = 7.2
Mim.Report.TextColor = RGB(192, 0, 0)  '= red!
Mim.Report.TextBold = 0
Mim.Report.TextItalic = 0
Mim.Report.TextUnderline = 0
Ktrl = Mim.Report.VPEPRINT(1, vsoftVAN, FULL_LINE)
Ktrl = Mim.Report.VPEPRINT(1, VsoftTot, FULL_LINE)
Return

CMDADRESSPACE:
Input #FlFree, adresXpos, adresYpos, adresXpos2, adresYpos2 ', adresBox
Ktrl = Mim.Report.WriteBox(adresXpos, adresYpos, adresXpos2, adresYpos2, "")
Return

CMDWRITE:
Input #FlFree, xPos, YPOS, xPos2, YPOS2, pdfFontSize, pdfFontName, pdfColor, pdfAlign, pdfFontBold, pdfFontItalic, pdfFontUnderLine
Line Input #FlFree, textstring
Do
    Line Input #FlFree, texttmp
    If texttmp = "CMD-ENDWRITE" Then
        Exit Do
    Else
        textstring = textstring & vbCrLf & texttmp
    End If
Loop
Mim.Report.FontName = pdfFontName
Mim.Report.FontSize = pdfFontSize
Mim.Report.TextColor = pdfColor
Mim.Report.TextBold = pdfFontBold
Mim.Report.TextItalic = pdfFontItalic
Mim.Report.TextUnderline = pdfFontUnderLine

Mim.Report.TextAlignment = pdfAlign
Ktrl = Mim.Report.Write(xPos, YPOS, xPos2, YPOS2, textstring)
Return
           
CMDWRITEBOX:
Input #FlFree, xPos, YPOS, xPos2, YPOS2, pdfFontSize, pdfFontName, pdfColor, pdfAlign, pdfFontBold, pdfFontItalic, pdfFontUnderLine
Line Input #FlFree, textstring
Do
    Line Input #FlFree, texttmp
    If texttmp = "CMD-ENDWRITE" Then
        Exit Do
    Else
        textstring = textstring & vbCrLf & texttmp
    End If
Loop
Mim.Report.FontName = pdfFontName
Mim.Report.FontSize = pdfFontSize
Mim.Report.TextColor = pdfColor
Mim.Report.TextBold = pdfFontBold
Mim.Report.TextItalic = pdfFontItalic
Mim.Report.TextUnderline = pdfFontUnderLine

Mim.Report.TextAlignment = pdfAlign
Ktrl = Mim.Report.WriteBox(xPos, YPOS, xPos2, YPOS2, textstring)
Return
           
CMDPRINT:
Input #FlFree, xPos, YPOS, pdfFontSize, pdfFontName, pdfColor, pdfFontBold, pdfFontItalic, pdfFontUnderLine
Line Input #FlFree, textstring
Mim.Report.FontName = pdfFontName
Mim.Report.FontSize = pdfFontSize
Mim.Report.TextColor = pdfColor
Mim.Report.TextBold = pdfFontBold
Mim.Report.TextItalic = pdfFontItalic
Mim.Report.TextUnderline = pdfFontUnderLine
Ktrl = Mim.Report.VPEPRINT(xPos, YPOS, textstring)
Return
      
End Sub

Private Sub mTaal_Click(Index As Integer)

For T = 1 To 4
    mTaal(T).Checked = False
Next

mTaal(Index).Checked = True
TaalKode = Index

End Sub


Private Sub txtDemoTekst_Change()

    Me.Label1 = Me.txtDemoTekst
    
End Sub

Private Sub ViaKB_Click(Index As Integer)

Dim HetBESTAND As String

TRYAGAINKB:
Select Case Index
    Case 0
        HetBESTAND = LOCATION_COMPANYDATA + "vpeSjbs\pdfDDEF" + Format(KopVoet) + Format(TaalKode) + Format(DocuType) + ".Txt"
        If Dir(HetBESTAND) = "" Then
            MsgBox LOCATION_COMPANYDATA + "vpeSjbs\pdfDDEF" + Format(KopVoet) + Format(TaalKode) + Format(DocuType) + ".Txt niet gevonden in de bedrijfsinhoudsopgave. Hierna wordt een voorbeelddocument vanuit de programmainhoudsopgave gecopiëerd", vbInformation
            If Not CopyFile(App.path & "\VpeSjbs", LOCATION_COMPANYDATA + "vpeSjbs\", "pdfDDEF" + Format(KopVoet) + Format(TaalKode) + Format(DocuType) + ".Txt") Then
                MsgBox App.path & "\VpeSjbs\pdfDDEF" + Format(KopVoet) + Format(TaalKode) + Format(DocuType) + ".Txt kan als voorbeelddocument niet vanuit de programmainhoudsopgave gecopiëerd worden.  Probeer eventueel manueel", vbInformation
                Exit Sub
            Else
                Ktrl = CopyFile(App.path & "\VpeSjbs", LOCATION_COMPANYDATA, "demo-compagny-logo.bmp")
                GoTo TRYAGAINKB
            End If
        Else
            Shell "notepad.exe " + HetBESTAND, vbNormalFocus
        End If
    Case 1
        HetBESTAND = LOCATION_COMPANYDATA + "vpeSjbs\pdfDDEF" + Format(KopVoet) + Format(TaalKode) + Format(DocuType) + "G.Txt"
        If Dir(HetBESTAND) = "" Then
            MsgBox HetBESTAND & " is niet aanwezig", vbInformation
        Else
            Shell "notepad.exe " + HetBESTAND, vbNormalFocus
        End If
End Select

End Sub
