VERSION 5.00
Begin VB.Form Vsoft 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Info Licentie"
   ClientHeight    =   5265
   ClientLeft      =   435
   ClientTop       =   1110
   ClientWidth     =   12675
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "0VSOFTMK.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5265
   ScaleWidth      =   12675
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdLeesMij 
      BackColor       =   &H00FFFFFF&
      Caption         =   "LeesMij"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   7
      Top             =   4680
      Width           =   2175
   End
   Begin VB.CommandButton cmdLog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Log"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton Ok 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Cancel          =   -1  'True
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   4680
      Width           =   855
   End
   Begin VB.Label AppInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label5"
      ForeColor       =   &H00FF0000&
      Height          =   645
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   3120
      Width           =   4545
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblCloud 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "marIntegraal 2026"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   840
      TabIndex        =   8
      Top             =   3840
      Width           =   4425
      WordWrap        =   -1  'True
   End
   Begin VB.Label AppInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label4"
      ForeColor       =   &H00FF0000&
      Height          =   945
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   2040
      Width           =   4455
      WordWrap        =   -1  'True
   End
   Begin VB.Label AppInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label2"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   5760
      TabIndex        =   3
      Top             =   4800
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   4455
      Left            =   5880
      Picture         =   "0VSOFTMK.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6825
   End
   Begin VB.Label LabelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Produktnaam"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   675
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   5625
   End
   Begin VB.Label LabelInfo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1035
      Index           =   2
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   4365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   645
      Left            =   6960
      TabIndex        =   9
      Top             =   3960
      Width           =   5625
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Vsoft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long


Dim Naam As String * 40
Dim BTWNummer As String * 12
Dim PeriodeKontrakt As String * 23
Dim TypeKontrakt As String * 17
Dim TempoLokatie As String
Dim EindeKontrakt As Integer



Private Sub cmdLeesMij_Click()

    Dim r As Long
    
    Dim ZoekHTM As Variant
    'ZoekHTM = App.path & "\LeesMij" & App.Major & App.Minor & Format(App.Revision, "000") & ".htm"
    ZoekHTM = "https://rv.be/accounting"
    'If Dir(ZoekHTM) <> "" Then
        Err = 0
        On Error Resume Next
        r = ShellExecute(0, "open", ZoekHTM, 0, 0, 1)
        DoEvents
        Mim.Arrange 0
        Ok_Click
    'End If
    
End Sub

Private Sub cmdLog_Click()

    On Error Resume Next
    If Dir(App.path + "\" + App.EXEName + ".log") = "" Then
        MsgBox "Enkel deze knop te gebruiken om LOGBOEK te kontroleren..."
    Else
        If Shell("notepad " + App.path + "\" + App.EXEName + ".log", 1) Then
        End If
    End If

End Sub


Private Sub Form_Activate()

    If Mim.Basis(3).Enabled = True Then
        cmdLog.Enabled = True
    Else
        cmdLog.Enabled = False
    End If

End Sub


Private Sub Form_Click()

    Unload Me
    
End Sub

Private Sub Form_DblClick()

    Unload Me
    
End Sub

Private Sub Form_Load()

Dim GeldigVanTot As String
Dim GeldigVan As String
Dim GeldigTot As String
Dim XX As Integer

Dim cnnLic As ADODB.Connection
Dim rsLic As ADODB.Recordset

LabelInfo(0).Caption = App.ProductName & "-{LocalFirst}"
AppInfo(1).Caption = App.LegalCopyright
AppInfo(2).Caption = App.FileDescription
AppInfo(0).Caption = Mid(appTitleAndVersion, InStr(appTitleAndVersion, "v.")) '"Versie " + Format(App.major) + "." + Format(App.minor) + "." + Format(App.Revision)

MousePointer = vbHourglass
Set NTRuimte = Workspaces(0)

Set cnnLic = New ADODB.Connection
Set rsLic = New ADODB.Recordset

BAModus = 1
'XX = InitBestanden

Dim sVar1 As String
Dim sVar2 As String
Dim sVar3 As String
Dim sVar4 As String
Dim sVar5 As String

cnnLic.Open ADOJET_PROVIDER & "Data Source=" & App.path & "\marnt.mdv;Persist Security Info=False"
'rsLic.Open "SELECT * FROM Allerlei WHERE v005 = 'zzzzzzzzzzzzzzzzzzzz'", cnnLic, adOpenForwardOnly, adLockReadOnly
rsLic.Open "SELECT TOP 1 * FROM Allerlei", cnnLic, adOpenForwardOnly, adLockOptimistic
If Ktrl Then
    MsgBox "Onverwachte situatie bij licentiecontrole"
    End
Else
    'sVar1 = Trim$(ADOBIB_TEXT(rsLic("MEMO"), "#vx01 #"))
    sVar1 = "Servicejaar 2026"
    ProgrammaVersie = sVar1
    sVar2 = sVar1
          
          Msg = sVar1 + vbCrLf
    sVar3 = Trim$(ADOBIB_TEXT(rsLic("MEMO"), "#A100 #"))
    sVar3 = "Pré 1990 gebruikers"
    'Msg = Msg + "(Bedrijf)" + vbCrLf + sVar3 + vbCrLf
    
    sVar4 = Trim$(ADOBIB_TEXT(rsLic("MEMO"), "#v008 #"))
    sVar4 = "000-000-000"
    'Msg = Msg + "(BTW- of NAT.Nr)" + vbCrLf + sVar4 + vbCrLf
    
    'sVar5 = Trim$(ADOBIB_TEXT(rsLic("MEMO"), "#vx02 #"))
    sVar5 = "01/12/2024 - 31/01/2030"
    
    Msg = Msg + "(Geldig van/tot)" + vbCrLf + sVar5
        
    LabelInfo(2).Caption = Msg
End If

GeldigVanTot = Left(sVar5, 10)
GeldigVan = Mid(GeldigVanTot, 7, 4) + Mid(GeldigVanTot, 4, 2) + Mid(GeldigVanTot, 1, 2)
GeldigVanTot = Right(sVar5, 10)
GeldigTot = Mid(GeldigVanTot, 7, 4) + Mid(GeldigVanTot, 4, 2) + Mid(GeldigVanTot, 1, 2)
If Format(Now, "yyyymmdd") < GeldigVan Or Format(Now, "yyyymmdd") > GeldigTot Then
    MsgBox "Datum van vandaag valt buiten de geldigheidsduur van uw contract !" & vbCr & "Voor meer inlichtingen mail uitsluitend naar info@rv.be" & vbCr & vbCr & "Het programma kan hierna verder gebruikt worden met beperkte functionaliteit", vbExclamation, "Controle van licentie !"
    usrLicentieInfo = "DemoModus, programma draait buiten contract.  Ga naar www.rv.be en open een helpticket of mail naar info@rv.be voor inlichtingen"
End If

Dim WinFlags As Long
Dim Mode As String, Processor As String

rsLic.Close
cnnLic.Close
Set rsLic = Nothing
Set cnnLic = Nothing
MousePointer = vbNormal

End Sub

Private Sub Image1_Click()

    Unload Me

End Sub

Private Sub LabelInfo_Click(Index As Integer)

    Unload Me

End Sub

Private Sub Ok_Click()

If EindeKontrakt = True Then
    End
End If
Unload Me

End Sub


