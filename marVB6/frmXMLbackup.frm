VERSION 5.00
Begin VB.Form frmXMLbackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XML Backup maken"
   ClientHeight    =   3870
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   8910
   Begin VB.CheckBox cbXMLopkuis 
      Alignment       =   1  'Right Justify
      Caption         =   "Na backup XML bestanden automatisch verwijderen"
      Height          =   555
      Left            =   6480
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1680
      Value           =   1  'Checked
      Width           =   2355
   End
   Begin VB.CommandButton cmdRestore 
      Caption         =   "Restore"
      Height          =   375
      Left            =   7140
      TabIndex        =   4
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CheckBox ckOverschrijven 
      Alignment       =   1  'Right Justify
      Caption         =   "steeds &Overschrijven"
      Height          =   315
      Left            =   7020
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1020
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.ListBox lbVerloop 
      Height          =   3765
      ItemData        =   "frmXMLbackup.frx":0000
      Left            =   0
      List            =   "frmXMLbackup.frx":0002
      TabIndex        =   1
      Top             =   0
      Width           =   6315
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Backup"
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   60
      Width           =   1695
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   7200
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   540
      Width           =   1695
   End
   Begin VB.Label lblRecords 
      Alignment       =   1  'Right Justify
      Height          =   195
      Left            =   7740
      TabIndex        =   6
      Top             =   2820
      Width           =   1035
   End
End
Attribute VB_Name = "frmXMLbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit




Private Sub cmdSluiten_Click()

    Unload Me
    
End Sub

Private Sub cmdStart_Click()

    Dim TSQsTring As String
    Dim sTBLName As String
    Dim strMario As String
        strMario = "C:\marnt\mario\"
    
    TSQsTring = OpenSchemaString("TABLE")

    'Overloop de tabellen
    Dim adoRShier As ADODB.Recordset
    Set adoRShier = New ADODB.Recordset
    
    '
    lbVerloop.Clear
    adoRShier.CursorLocation = adUseClient
    Do While TSQsTring <> ""
        sTBLName = Mid(TSQsTring, 1, InStr(TSQsTring, vbCr) - 1)
        TSQsTring = Mid(TSQsTring, InStr(TSQsTring, vbCr) + 1)
        Screen.MousePointer = vbHourglass
        lbVerloop.AddItem Now & " Start ophalen tabel: " & sTBLName
        Me.Refresh
        adoRShier.Open "SELECT * FROM " & sTBLName, adntDB
        lbVerloop.AddItem Now & " Einde ophalen tabel: " & sTBLName
        Me.Refresh
        lbVerloop.AddItem Now & " Start wegschrijven: marNET" & sTBLName & ".xml"
        Me.Refresh
        If ckOverschrijven.Value = vbChecked Then
            If Dir(strMario & "marNET" & sTBLName & ".xml") <> "" Then
                Kill strMario & "marNET" & sTBLName & ".xml"
            End If
            adoRShier.Save strMario & "marNET" & sTBLName & ".xml", adPersistXML
        ElseIf Dir(strMario & "marNET" & sTBLName & ".xml") <> "" Then
            If MsgBox("Bestaande " & "marNET" & sTBLName & ".xml overschrijven.  Bent U zeker", vbYesNo + vbQuestion + vbDefaultButton1) = vbYes Then
                Kill strMario & "marNET" & sTBLName & ".xml"
                adoRShier.Save strMario & "marNET" & sTBLName & ".xml", adPersistXML
            End If
        End If
        Me.Refresh
        adoRShier.Close
    Loop
    Set adoRShier = Nothing
    MsgBox "Einde XML backup.  Bij voorkeur vervolgens de xml bestanden met uw favoriet compressieprogramma comprimeren naar één kleiner bestand." & vbCr & vbCr & "De XML bestanden bevinden zich in de HOOFDROOT van uw programmmainhoudsopgave:" & vbCr & strMario
    Screen.MousePointer = vbNormal
        
End Sub


Function UnZipHet()
  
  Dim MijnToepassingID
  
  ' AppActivate kan eveneens gebruikt worden via ID bekomen van de Shell functie.
    On Error Resume Next
    MijnToepassingID = Shell(Left(ProgrammaLokatie, 3) & "Program Files\WinZip\WINZIP32.EXE -e " & Left(ProgrammaLokatie, 3) & "marXML" & Mid(BedrijfsLokatie, Len(BedrijfsLokatie) - 3, 3) & ".zip " & Left(ProgrammaLokatie, 3), vbNormalFocus) ' & "marNET*.xml"
    DoEvents
    AppActivate MijnToepassingID, True
    DoEvents
    
End Function

