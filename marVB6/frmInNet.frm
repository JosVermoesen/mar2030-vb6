VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmMarNtInetUpgrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "marIntegraal NT upgraden via Internet !"
   ClientHeight    =   3735
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8865
   Begin VB.CommandButton cbStoppen 
      Caption         =   "&Stoppen"
      Height          =   300
      Left            =   7260
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1445
   End
   Begin VB.ComboBox cbKeuze 
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   600
      Width           =   7335
   End
   Begin VB.ComboBox cmbITType 
      Height          =   315
      Left            =   1440
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   60
      Width           =   3675
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   300
      Left            =   7260
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1445
   End
   Begin MSComCtl2.Animation aniVIDEO 
      Height          =   405
      Left            =   7680
      TabIndex        =   9
      Top             =   2160
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   714
      _Version        =   393216
      Center          =   -1  'True
      FullWidth       =   45
      FullHeight      =   27
   End
   Begin VB.CommandButton cmdGET 
      Caption         =   "&GET !"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   7260
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   960
      Width           =   1445
   End
   Begin VB.CommandButton cmdDIR 
      Caption         =   "&Dir"
      Height          =   300
      Left            =   7260
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1740
      Width           =   1445
   End
   Begin VB.CommandButton cmdChangeDirectory 
      Caption         =   "&CD"
      Height          =   300
      Left            =   7260
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1380
      Width           =   1445
   End
   Begin VB.TextBox txtResponse 
      Height          =   2265
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      Top             =   1320
      Width           =   7095
   End
   Begin VB.TextBox txtRemotepath 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "/Software/*.*"
      Top             =   960
      Width           =   5685
   End
   Begin VB.TextBox txtURL 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   5160
      TabIndex        =   4
      Text            =   "ftp://sky39367:rv570428@users.skynet.be"
      ToolTipText     =   "voorbeeld : ftp://sky99999:paswoord@users.skynet.be"
      Top             =   60
      Width           =   3585
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6900
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bewaren in "
      Height          =   315
      Index           =   0
      Left            =   60
      TabIndex        =   11
      Top             =   600
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Verbindings- Parameters"
      Height          =   495
      Index           =   1
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   1350
   End
   Begin VB.Label LblParameter 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "&Bevel/Resultaat"
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   1365
   End
End
Attribute VB_Name = "frmMarNtInetUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cbStoppen_Click()

    Inet1.Cancel
    
End Sub

Private Sub cmbITType_Click()

Inet1.Protocol = Left(cmbITType.Text, 1)
If txtURL.Text = "ftp://sky39367:rv570428@users.skynet.be" Then
    txtURL.PasswordChar = "*"
Else
    txtURL.PasswordChar = ""
End If

End Sub

Private Sub cmdChangeDirectory_Click()

On Error Resume Next
Err = 0
' Change directory to txtRemotePath.
Inet1.Execute txtURL.Text, "CD " & txtRemotepath.Text
If Err Then
    txtResponse = Error
End If

End Sub

Private Sub cmdSluiten_Click()

Unload Me

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Dim rtPath As String * 260
Dim PathHier As String
Dim Ret As Integer
Dim TelTot As Integer
    
For TelTot = 0 To 100
    On Error Resume Next
    rtPath = ""
    Err = 0
    Ret = SHGetFolderPath(0, TelTot, 0, 0, rtPath)
    If Err Then
    Else
        PathHier = Mid(rtPath, 1, InStr(rtPath, Chr(0)) - 1)
        cbKeuze.AddItem PathHier
    End If
Next
cbKeuze.AddItem "c:", 0
cbKeuze.ListIndex = 0

cmbITType.AddItem "2: FTP (File Transfer Protocol)"
cmbITType.AddItem "4: HTTP (HyperText Transfer Protocol)"
cmbITType.AddItem "5: HTTPS (Secure HTTP)"
cmbITType.ListIndex = 0

Me.Left = 0
Me.Top = 0

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
' Retrieve server response using the GetChunk
' method when State = 12.
Dim vtData As Variant    ' Data variable.
Dim bData() As Byte      ' Data variable

'TEDOEN: controleren of dit niet beter is !
'Private Function GetData()
'Dim vData As Variant
'Dim strTemp As String
'Do
'  vData = Inet1.GetChunk(256, icString)
'  DoEvents
'  If Len(vData) Then
'    strTemp = strTemp & vData
'    lblBytesReceived.Caption = CStr$(Len(strTemp))
'  End If
'Loop While Len(vData)
'GetData = strTemp
'End Function

Select Case State
        ' ... Other cases not shown.
    Case icError
        ' 11
        ' In case of error, return ResponseCode and
        ' ResponseInfo.
        txtResponse = Inet1.ResponseCode & ":" & Inet1.ResponseInfo
        Mim.MousePointer = vbNormal
        Mim.WindowState = vbNormal
        
        aniVIDEO.Close
        aniVIDEO.Visible = False
        aniVIDEO.AutoPlay = False
        Mim.WindowState = vbNormal
        Mim.MousePointer = vbNormal
        
    Case icResponseCompleted
        ' 12
        
        txtResponse = Inet1.ResponseCode & ":" & Inet1.ResponseInfo
        
        aniVIDEO.Close
        aniVIDEO.Visible = False
        aniVIDEO.AutoPlay = False
        Mim.WindowState = vbNormal
        Mim.MousePointer = vbNormal
                
        Dim strData As String
        Dim bDone As Boolean
        bDone = False
        ' Get first chunk.
        vtData = Inet1.GetChunk(1024, icString)
        DoEvents
        Do While Not bDone
            strData = strData & vtData
            ' Get next chunk.
            vtData = Inet1.GetChunk(1024, icString)
            DoEvents
            If Len(vtData) = 0 Then
                bDone = True
            End If
        Loop
        txtResponse = strData + vbCrLf + "Opdracht voltooid"
                
    Case Else
        txtResponse = Inet1.ResponseCode & ":" & Inet1.ResponseInfo
End Select

End Sub


Private Sub cmdDIR_Click()

On Error Resume Next
Err = 0
Inet1.Execute txtURL.Text, "DIR " & txtRemotepath.Text
If Err Then
    txtResponse = Error
End If

End Sub

Private Sub cmdGET_Click()

If Dir(cbKeuze.Text & "\" & txtRemotepath.Text) <> "" Then
    KtrlBox = MsgBox(cbKeuze.Text & "\" & txtRemotepath.Text & " zal overschreven worden !  Akkoord ?", vbYesNo + vbDefaultButton2 + vbQuestion)
    If KtrlBox = vbYes Then
        Kill cbKeuze.Text & "\" & txtRemotepath.Text
    Else
        Exit Sub
    End If
End If

If txtRemotepath.Text = "marintegraal.exe" Then
    Msg = "Hierna wordt het bestand " & txtRemotepath.Text & " opgehaald en geplaatst in " & cbKeuze & "\" & vbCr & vbCr
    Msg = Msg & "marIntegraal Studio wordt geMinimaliseerd totdat de overdracht beëindigd is.  U mag uiteraard verderwerken met het FTP scherm open.  U vertraagt enkel wel het ophalen, zodat dit niet aanbevolen is voor tragere modemverbindingen (niet ISDN verbindingen)"
    MsgBox Msg, vbInformation
    Mim.WindowState = vbMinimized
End If

Mim.MousePointer = vbHourglass
aniVIDEO.Open App.path + "\download.avi"
aniVIDEO.Visible = True
aniVIDEO.AutoPlay = True

On Error Resume Next
Err = 0

Inet1.Execute txtURL.Text, "GET " + txtRemotepath.Text + " " + cbKeuze.Text + "\" + txtRemotepath.Text
If Err Then
    txtResponse = Error
End If

End Sub


Private Sub txtURL_GotFocus()

If txtURL.Text = "ftp://sky39367:rv570428@users.skynet.be" Then
    Beep
    SnelHelpPrint "R&Vsoftware FTP paswoord wordt nu verwijderd.  U kan vervolgens zelf de nodige paswoorden/adresgegevens inbrengen voor toegang tot Provider en/of systemen", blLogging
    txtURL.Text = ""
    txtURL.PasswordChar = ""
End If

End Sub
