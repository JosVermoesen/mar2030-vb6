VERSION 5.00
Begin VB.Form SetupEnParameters 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup Boekjaar En Parameters"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5205
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   HelpContextID   =   21000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2310
   ScaleWidth      =   5205
   Begin VB.OptionButton SetupOption 
      Caption         =   "Kassaverkoop"
      Height          =   240
      Index           =   8
      Left            =   60
      TabIndex        =   10
      Top             =   2010
      Width           =   3555
   End
   Begin VB.OptionButton SetupOption 
      Caption         =   "Status Boekjaar"
      Height          =   240
      Index           =   7
      Left            =   60
      TabIndex        =   9
      Top             =   1770
      Width           =   3555
   End
   Begin VB.OptionButton SetupOption 
      Caption         =   "Bedrijfsinformatie"
      Height          =   240
      Index           =   6
      Left            =   60
      TabIndex        =   8
      Top             =   1530
      Value           =   -1  'True
      Width           =   3555
   End
   Begin VB.OptionButton SetupOption 
      Caption         =   "Financieel en Rekeningen"
      Height          =   240
      Index           =   5
      Left            =   60
      TabIndex        =   7
      Top             =   1290
      Width           =   3555
   End
   Begin VB.OptionButton SetupOption 
      Caption         =   "Default Collectieve Rekeningen"
      Height          =   240
      Index           =   4
      Left            =   60
      TabIndex        =   6
      Top             =   1050
      Width           =   3555
   End
   Begin VB.OptionButton SetupOption 
      Caption         =   "BTW Default Rekeningen"
      Height          =   240
      Index           =   3
      Left            =   60
      TabIndex        =   5
      Top             =   810
      Width           =   3555
   End
   Begin VB.OptionButton SetupOption 
      Caption         =   "Verkoopverrichtingen"
      Height          =   240
      Index           =   2
      Left            =   60
      TabIndex        =   4
      Top             =   570
      Width           =   3555
   End
   Begin VB.OptionButton SetupOption 
      Caption         =   "Aankoopverrichtingen"
      Height          =   240
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   330
      Width           =   3555
   End
   Begin VB.OptionButton SetupOption 
      Caption         =   "Boekingen en algemene instellingen"
      Height          =   240
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   90
      Width           =   3555
   End
   Begin VB.CommandButton Afsluiten 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   320
      Left            =   3840
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   450
      Width           =   1240
   End
   Begin VB.CommandButton Wijzigen 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   320
      Left            =   3840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   90
      Width           =   1240
   End
End
Attribute VB_Name = "SetupEnParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Option Explicit

Private Sub Afsluiten_Click()

Unload Me

End Sub

Private Sub Form_Load()

If Not Toegankelijk(Me) Then
    Unload Me
    Exit Sub
End If

Top = 0
Left = 0

End Sub


Private Function TeleBibDEF(TeZoeken As String)
Dim FlInput         As Integer
Dim T               As Integer
Dim LokatieString   As String

TeleBibDEF = False

On Local Error GoTo TeleBibError

If Dir$(PROGRAM_LOCATION + "Def\" + "099.Def") = "" Then
    MsgBox "Geen TeleBib definitie 099.Def"
    Exit Function
End If

FlInput = FreeFile
Open PROGRAM_LOCATION + "Def\" + "099.Def" For Input As FlInput
Do While Not EOF(FlInput)
    Line Input #FlInput, LokatieString
    If LokatieString = TeZoeken Then
        Exit Do
    End If
Loop
If LokatieString <> TeZoeken Then
    Exit Function
End If

T = 0
Do While Not EOF(FlInput)
    Input #FlInput, TELEBIB_CODE(T), TELEBIB_TEXT(T), TELEBIB_TYPE(T), TELEBIB_LENGTH(T)
    If Left(TELEBIB_CODE(T), 1) = ";" Then
        Exit Do
    Else
        T = T + 1
    End If
Loop
Close FlInput
TELEBIB_CODE(T) = ""
TeleBibDEF = True
Exit Function

TeleBibError:
MsgBox "VSBibinlaadfout" + Str$(T) + " error:" + Error$
Close FlInput
TeleBibDEF = False
Exit Function
Resume

End Function

Private Sub SetupOption_DblClick(Index As Integer)

Wijzigen_Click

End Sub

Private Sub Wijzigen_Click()
Dim T       As Integer
Dim aa      As String
Dim CrText  As String
Dim CrText2 As String
Dim BoxMask As String
Dim BoxType As Integer

For COUNT_TO = 0 To 8
    If SetupOption(COUNT_TO).Value = True Then
        Exit For
    End If
Next

If TeleBibDEF(";" + SetupOption(COUNT_TO).Caption) Then
Else
    Beep
    MsgBox "Definitiebestand 099.DEF defekt ?"
    Exit Sub
End If

Unload Xlog
Load Xlog
Xlog.Height = Xlog.Height + 3
Xlog.X.Rows = 1
Xlog.X.Cols = 3

aa = ""
T = 0
Do While TELEBIB_CODE(T) <> Space$(10)
    bGet TABLE_COUNTERS, 0, Mid(TELEBIB_CODE(T), 5, 5)
    If Ktrl Then
        CrText = ""
    Else
        RecordToVeld TABLE_COUNTERS
        'CrText = ntRS(TABLE_COUNTERS).Fields("v217")
        CrText = rsMAR(TABLE_COUNTERS).Fields("v217")
        Select Case Mid(TELEBIB_CODE(T), 2, 2)
            Case "  ", "K ", "L ", "LC", "R ", "R3", "R4", "R6", "R7"
                'niks
            Case Else
                Select Case Mid(TELEBIB_CODE(T), 1, 1)
                    Case " "
                        BoxMask = "00"
                        BoxType = 0
                    Case "0" To "9"
                        BoxMask = "000"
                        BoxType = 1
                End Select
                If Left(TELEBIB_CODE(T), 1) = "@" Or CrText = "" Then
                Else
                    CrText = fmarBoxText(Format(Val(Mid(TELEBIB_CODE(T), 1, 3)), BoxMask), "2", CrText)    'hier eventueel taaloptie
                End If
        End Select
    End If
    aa = TELEBIB_CODE(T) & vbTab & TELEBIB_TEXT(T) & vbTab & CrText
    T = T + 1
    Xlog.X.AddItem aa
Loop

Xlog.X.Col = 0
Xlog.X.Row = 0
Xlog.X.Text = "vsfKode"
Xlog.X.Col = 1
Xlog.X.Text = "Veldomschrijving"
Xlog.X.Col = 2
Xlog.X.Text = "Veldgegevens"

Xlog.X.Row = 1
Xlog.X.Col = 0

Xlog.X.ColWidth(0) = 960
Xlog.X.ColWidth(1) = 3405
Xlog.X.ColWidth(2) = 4395

Xlog.X.ColAlignment(0) = flexAlignLeftTop
Xlog.X.ColAlignment(1) = flexAlignLeftTop
Xlog.X.ColAlignment(2) = flexAlignLeftTop

XLogShow:
Xlog.WijzigenLijn.Default = True
Xlog.Afsluiten.TabStop = True
XLogKey = ""
Xlog.SSTab1.TabVisible(1) = False
Xlog.Show 1
If XLogKey <> "" Then
    Msg = "Boekjaarparameters overschrijven.  Bent U zeker ?"
    KtrlBox = MsgBox(Msg, 292)
    If KtrlBox = 6 Then
        T = 0
        Xlog.X.Col = 2
        Do While TELEBIB_CODE(T) <> Space$(10)
            Xlog.X.Row = T + 1
            CrText2 = Xlog.X.Text
            bGet TABLE_COUNTERS, 0, Mid(TELEBIB_CODE(T), 5, 5)
            If Ktrl Then
                TLB_RECORD(TABLE_COUNTERS) = ""
                vBib TABLE_COUNTERS, Mid(TELEBIB_CODE(T), 5, 5), "v071"
            Else
                RecordToVeld TABLE_COUNTERS
            End If
            
            Select Case Mid(TELEBIB_CODE(T), 2, 2)
                Case "  "
                Case Else
                    If Left(TELEBIB_CODE(T), 1) = "@" Then
                    Else
                        On Local Error Resume Next
                        CrText2 = Left(CrText2, InStr(CrText2, ":") - 1)
                        On Local Error GoTo 0
                    End If
            End Select
            vBib TABLE_COUNTERS, CrText2, "v217"
            If Ktrl Then
                bInsert TABLE_COUNTERS, 0
            Else
                bUpdate TABLE_COUNTERS, 0
            End If
            T = T + 1
        Loop
        Afsluiten.SetFocus
    End If
End If

End Sub

