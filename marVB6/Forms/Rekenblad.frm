VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "mschrt20.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   3780
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   3240
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdOverlopen 
      Caption         =   "Volgende"
      Height          =   375
      Index           =   3
      Left            =   4380
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOverlopen 
      Caption         =   "Vorige"
      Height          =   375
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOverlopen 
      Caption         =   "Laatste"
      Height          =   375
      Index           =   1
      Left            =   5640
      TabIndex        =   3
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdOverlopen 
      Caption         =   "Eerste"
      Height          =   375
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   3840
      Width           =   1215
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3555
      Left            =   120
      OleObjectBlob   =   "Rekenblad.frx":0000
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Width           =   6795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   5700
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsHier As ADODB.Recordset


Private Sub cmdOverlopen_Click(Index As Integer)

    Select Case Index
        Case 0
            rsHier.MoveFirst
            
        Case 1
            rsHier.MoveLast
            
        Case 2
            rsHier.MovePrevious
            If rsHier.BOF = True Then
                rsHier.MoveFirst
            End If
        
        Case 3
            rsHier.MoveNext
            If rsHier.EOF = True Then
                rsHier.MoveLast
            End If
            
    End Select
    Command1_Click
    
End Sub

Private Sub Command1_Click()
     
    Dim arrValues(1 To 10, 1 To 2)
    Dim i As Integer
    For i = 1 To 10
        arrValues(i, 1) = "Bj." & BJPERDAT.Boekjaar.List(i - 1)  ' Labels
        arrValues(i, 2) = rsHier(1 + i) ' Series 1 values.
        'arrValues(i, 3) = 2 * i ' Series 2 values.
    Next i
    MSChart1.TitleText = Trim$(rsHier(0)) & " " & Trim$(rsHier(1))
    MSChart1.ChartData = arrValues

End Sub

Private Sub Command2_Click()

    'Dim frmB As frmBrowser
    'Dim rvNl As String
    
    'rvNl = "%0d%0a"
    'Set frmB = New frmBrowser
    'If Err Then MsgBox Error: Exit Sub
    'frmB.StartingAddress = "mailto:" & _
        "jos@rv.be;jos.vermoesen@skynet.be?" & _
        "subject=een test onderwerp" & _
        "&body=Dit is een eerste lijn" & rvNl & _
            "Dit is een tweede lijn" & rvNl & _
            "Dit is een derde lijn"
            
    'frmB.Visible = False
    'If Err Then MsgBox Error:  Exit Sub
    'frmb.Show
    'If Err Then MsgBox Error:  Exit Sub
        
End Sub

Private Sub Command3_Click()

    'Dim Result As Boolean
    'Result = adxKolom("Klanten", "autoNum", adCurrency, 0)
        
End Sub

Private Sub Form_Load()
  
    Set rsHier = New ADODB.Recordset
    Screen.MousePointer = vbHourglass
    rsHier.CursorLocation = adUseClient
    rsHier.Open "SELECT v019, v020, dece022, dece023, dece024, dece025, dece026, dece027, dece028, dece029, dece030, dece031 FROM Rekeningen ORDER BY v019", adntDB, adOpenForwardOnly, adLockReadOnly
    Screen.MousePointer = vbNormal
    
End Sub
