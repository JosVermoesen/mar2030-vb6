VERSION 5.00
Object = "{0002E558-0000-0000-C000-000000000046}#1.1#0"; "owc11.dll"
Begin VB.Form frmRekenBlad 
   Caption         =   "marIntegraal.NET XP-Rekenblad"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13380
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7155
   ScaleWidth      =   13380
   Begin OWC11.Spreadsheet Spreadsheet1 
      Height          =   4335
      Left            =   2400
      OleObjectBlob   =   "marRekenBlad.frx":0000
      TabIndex        =   10
      Top             =   1440
      Width           =   8655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   11640
      TabIndex        =   9
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox cbXML 
      Caption         =   "XML filter"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   6840
      Value           =   1  'Checked
      Width           =   1035
   End
   Begin VB.CommandButton cmdSchoon 
      Caption         =   "Schoon"
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   6720
      Width           =   915
   End
   Begin VB.ComboBox cbBestandsNaam 
      Height          =   315
      Left            =   2100
      TabIndex        =   6
      Text            =   "cbBestandsNaam"
      ToolTipText     =   "bij ingave bestandnaam vb. test vergeet bij xml niet weg te schrijven als test.xml a.u.b.!"
      Top             =   6780
      Width           =   2955
   End
   Begin VB.TextBox txtNaam 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9240
      TabIndex        =   5
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtIdCode 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   8160
      TabIndex        =   4
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdZoeken 
      Caption         =   "&Zoeken"
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6720
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6720
      Width           =   915
   End
   Begin VB.CommandButton cmdBewaren 
      Caption         =   "&Bewaren"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6720
      Width           =   1035
   End
   Begin VB.CommandButton cmdInladen 
      Caption         =   "&Inladen"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6720
      Width           =   855
   End
End
Attribute VB_Name = "frmRekenBlad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fs      As FileSystemObject
Dim KtrlBox As Variant
Dim Msg     As Variant


Private Sub cbBestandsNaam_Click()

    DoEvents
    cmdInladen_Click
        
End Sub


Private Sub cbXML_Click()

    Dim strType As String
    Dim bstNaam As String
        
    If cbXML.Value = vbChecked Then
        strType = "xml"
    Else
        strType = "rvs"
    End If
    
    cbBestandsNaam.Clear
    bstNaam = Dir(LOCATION_COMPANYDATA & "xlsx-templates\*." & strType)
    If bstNaam = "" Then
        MsgBox "Er zijn (nog) geen ." & strType & " bestanden", vbInformation
    Else
        Do
            cbBestandsNaam.AddItem bstNaam
            bstNaam = Dir
            If bstNaam = "" Then
                Exit Do
            End If
        Loop
    End If
    If cbBestandsNaam.ListCount = 0 Then
    Else
        cmdSchoon_Click
        cbBestandsNaam.ListIndex = 0
    End If
    
End Sub

Private Sub cmdBewaren_Click()

    If MsgBox(cbBestandsNaam.Text & " bewaren.  Bent U zeker ?", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
        If UCase(Right(cbBestandsNaam.Text, 3)) = "XML" Then
            KtrlBox = ScrMaakXMLTekst(Spreadsheet1.XMLData, cbBestandsNaam.Text)
        Else
            KtrlBox = ScrMaakXMLTekst(Spreadsheet1.HTMLData, cbBestandsNaam.Text)
        End If
        If KtrlBox = False Then
            MsgBox "Data kan niet bewaard worden", vbExclamation
        Else
            MsgBox "Data met succes weggeschreven", vbInformation
        End If
    End If
    
End Sub

Private Sub cmdInladen_Click()

    Dim TekstZelf As Variant

    KtrlBox = ScrLeesXMLTekst(TekstZelf, cbBestandsNaam.Text, True)
    If KtrlBox = False Then
        MsgBox "Data kan niet ingeladen worden", vbExclamation
    Else
        'cbBestandsNaam.Enabled = False
    End If
    Form_Resize

End Sub

Private Sub cmdSchoon_Click()

    Spreadsheet1.HTMLData = ""
    cbBestandsNaam.Enabled = True
    
End Sub

Private Sub cmdSluiten_Click()
    
    Unload Me
    
End Sub

Private Sub cmdZoeken_Click()

    SharedFl = TABLE_CUSTOMERS
    aIndex = 1
    SqlSearch.Show 1
    If Ktrl Then
    Else
        txtIdCode.Text = rsMAR(TABLE_CUSTOMERS)("A110").Value
        txtNaam.Text = rsMAR(TABLE_CUSTOMERS)("A100").Value
        MsgBox "Stop"
        'RecordSamenStellen txtIdCode.Text, BOOKYEAR_FROMTO
    End If
    
End Sub

Private Sub Command1_Click()

    With Spreadsheet1
        .Range("A30").Value = 12
        .Range("A31").Value = 34
        .Range("A32").Formula = "=ATAN2(A30,A31)"
        .Range("A33").Formula = "=A32*180/PI()"
    End With

End Sub

Private Sub Form_Load()

    Dim bstNaam As String
    
    Set fs = New FileSystemObject
    With Spreadsheet1
        .TitleBar.Caption = "marIntegraal.NET Rekenblad"
        .DisplayTitleBar = False
        .DisplayToolbar = False
    End With
    
    bstNaam = Dir(LOCATION_COMPANYDATA & "xlsx-templates\*.xml")
    If bstNaam = "" Then
        MsgBox "Er zijn nog geen .xml bestanden", vbInformation
        
    Else
        Do
            cbBestandsNaam.AddItem bstNaam
            bstNaam = Dir
            If bstNaam = "" Then
                Exit Do
            End If
        Loop
    End If
    If cbBestandsNaam.ListCount <> 0 Then
        cbBestandsNaam.ListIndex = 0
        cmdInladen_Click
    End If
    WindowState = vbMaximized
    
    'With Spreadsheet1
    '    .Range("A1").Value = 12
    '    .Range("A2").Value = 34
    '    .Range("A3").Formula = "=ATAN2(A1,A2)"
    '    .Range("A4").Formula = "=A3*180/PI()"
    'End With
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    With Spreadsheet1
        .Left = 0
        .Top = 0
        .Height = Me.Height - 1200
        .Width = Me.Width - 100
    End With
    cmdBewaren.Top = Spreadsheet1.Height + 240
    cmdSchoon.Top = Spreadsheet1.Height + 240
    cmdZoeken.Top = Spreadsheet1.Height + 240
    cmdInladen.Top = Spreadsheet1.Height + 240
    cbXML.Top = Spreadsheet1.Height + 240
    cmdSluiten.Top = Spreadsheet1.Height + 240
    cbBestandsNaam.Top = Spreadsheet1.Height + 240
    txtIdCode.Top = Spreadsheet1.Height + 240
    txtNaam.Top = Spreadsheet1.Height + 240

End Sub

Function ScrMaakXMLTekst(DataZelf As Variant, Bestandsnaam As String) As Boolean

    On Error Resume Next
    Dim A As TextStream
    Screen.MousePointer = vbHourglass
    Set A = fs.CreateTextFile(LOCATION_COMPANYDATA & "xlsx-templates\" & Bestandsnaam, True)
    A.WriteLine (DataZelf)
    A.Close
    If Err Then
        ScrMaakXMLTekst = False
    Else
        ScrMaakXMLTekst = True
    End If
    A = Nothing
    Screen.MousePointer = vbNormal
    
End Function

Function ScrLeesXMLTekst(DataZelf As Variant, Bestandsnaam As String, LeesAlles As Boolean) As Boolean
    
    On Error Resume Next
    Const ForReading = 1
    Dim A As TextStream
        
    Set A = fs.OpenTextFile(LOCATION_COMPANYDATA & "xlsx-templates\" & Bestandsnaam, ForReading, False, TristateFalse)
    If Err Then ScrLeesXMLTekst = False: Exit Function
    
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
        ScrLeesXMLTekst = False
    Else
        ScrLeesXMLTekst = True
    End If
    A = Nothing
    
End Function


