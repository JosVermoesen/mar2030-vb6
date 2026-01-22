VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form BJPERDAT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datum / Periode / Boekjaar (Ctrl+D)"
   ClientHeight    =   1155
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   4440
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   HelpContextID   =   22000
   Icon            =   "00000000.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1155
   ScaleWidth      =   4440
   WindowState     =   1  'Minimized
   Begin MSComCtl2.DTPicker DatumVerwerking 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Top             =   60
      Width           =   1335
      _ExtentX        =   2355
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
      CurrentDate     =   43069
      MaxDate         =   58862
      MinDate         =   31048
   End
   Begin VB.ComboBox PeriodeBoekjaar 
      BackColor       =   &H00C0C0C0&
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
      Height          =   315
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   420
      Width           =   2895
   End
   Begin VB.ComboBox Boekjaar 
      BackColor       =   &H00C0C0C0&
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
      Height          =   315
      Left            =   3330
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   750
      Width           =   1065
   End
   Begin VB.CommandButton Verkleinen 
      Cancel          =   -1  'True
      Caption         =   "Minimaliseren"
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
      Height          =   350
      Left            =   3120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   60
      Width           =   1240
   End
   Begin VB.Label Label1 
      Caption         =   "&Datum vandaag"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Actieve &Periode"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   60
      TabIndex        =   2
      Top             =   450
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Actief &Boekjaar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   1860
      TabIndex        =   4
      Top             =   810
      Width           =   1380
   End
End
Attribute VB_Name = "BJPERDAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Private Sub Boekjaar_Click()
Dim A                       As String * 16
Dim XX                      As String
Dim YY                      As String
Dim FlTemp                  As Integer
Dim AktievePeriode          As Integer
Dim T                       As Integer
Dim PeriodeMeestLogisch     As Integer

PeriodeMeestLogisch = -1
If ACTIVE_BOOKYEAR <> Boekjaar.ListIndex Then
    bClose 99
    TABLEDEF_ONT(TABLE_COUNTERS) = Format(Boekjaar.ListIndex, "00") + ".ONT"
    bstNaam(TABLE_COUNTERS) = "jr" & Boekjaar.Text
    CloseOpenWindows
    ACTIVE_BOOKYEAR = Boekjaar.ListIndex
    AktievePeriode = 1
    PeriodeBoekjaar.Clear
    FlTemp = FreeFile
    Open LOCATION_COMPANYDATA + "DEF" + Format(ACTIVE_BOOKYEAR, "00") + ".OCT" For Random As FlTemp Len = 16
    For T = 1 To 99
        Get FlTemp, T, A
        If A = Space$(16) Then
                PeriodeBoekjaar.ListIndex = 0
                YY = PeriodeBoekjaar.Text
                BOOKYEAR_FROMTO = Mid(YY, 7, 4) + Mid(YY, 4, 2) + Mid(YY, 1, 2) + Mid(XX, 20, 4) + Mid(XX, 17, 2) + Mid(XX, 14, 2)
                Exit For
        Else
                'If PeriodeBoekjaar.ListCount = 0 Then
                '    XX = "00" + "/" + Mid(A, 5, 2) + "/" + Left(A, 4) + " - " + Right(A, 2) + "/" + Mid(A, 13, 2) + "/" + Mid(A, 9, 4)
                'Else
                XX = Mid(A, 7, 2) + "/" + Mid(A, 5, 2) + "/" + Left(A, 4) + " - " + Right(A, 2) + "/" + Mid(A, 13, 2) + "/" + Mid(A, 9, 4)
                'End If
                PeriodeBoekjaar.AddItem XX
                If DATE_KEY(MIM_GLOBAL_DATE) >= DATE_KEY(Left(XX, 10)) And DATE_KEY(MIM_GLOBAL_DATE) <= DATE_KEY(Right(XX, 10)) Then
                    PeriodeMeestLogisch = T - 1
                End If
        End If
    Next T
    If PeriodeMeestLogisch <> -1 Then
        PeriodeBoekjaar.ListIndex = PeriodeMeestLogisch
    Else
        PeriodeBoekjaar.ListIndex = AktievePeriode - 1
    End If
    Close FlTemp
    Close FL99
    FL99 = 0
End If
'If PeriodeBoekjaar.ListIndex < 0 Then
'    If PeriodeBoekjaar.ListCount Then
'        PeriodeBoekjaar.ListIndex = 0
'    End If
'End If
XX = PeriodeBoekjaar.Text
PERIOD_FROMTO = Mid(XX, 7, 4) + Mid(XX, 4, 2) + Left(XX, 2) + Right(XX, 4) + Mid(XX, 17, 2) + Mid(XX, 14, 2)

XisEuroWisBEF = False
If Mim.SnelHelp.Panels(2).Text = "---" Then
Else
    If Mim.SnelHelp.Panels(2).Text = "EUR" Then
        XisEuroWisBEF = True
    End If
    If Len(String99(READING, 296)) = 0 Then
        MsgBox "Gelieve Setup Boekingen en algemene instellingen : munt van de Boekhouding in te stellen a.u.b.  Pér bedrijf, pér boekjaar.  Hierna wordt voorlopig verder gewerkt in BEF."
        bhEuro = False
    ElseIf String99(READING, 296) = "BEF" Then
        bhEuro = False
    ElseIf String99(READING, 296) = "EUR" Then
        bhEuro = True
    Else
        SnelHelpPrint "Onlogische situatie", BL_LOGGING
        bhEuro = False
    End If
    If bhEuro Then
        XisEuroWisBEF = False
        Mim.SnelHelp.Panels(2).Text = "EUR"
    Else
        Mim.SnelHelp.Panels(2).Text = "BEF"
        If XisEuroWisBEF = True And Boekjaar.ListIndex = 1 Then
        Else
            XisEuroWisBEF = False
        End If
    End If
    SnelHelpPrint "XisEuroWisBEF = " & XisEuroWisBEF, BL_LOGGING
    If Boekjaar.ListIndex = 1 And bhEuro = False Then
        Msg = "Dit bedrijf heeft een boekjaar -1 met verwerking in BEF.  Indien U nog boekingen in BEF wenst uit te voeren voor het zopas aangeduide boekjaar, gelieve een vorige versie (6.5.300 of lager) opnieuw te installeren a.u.b."
        MsgBox Msg, vbExclamation
    End If
End If

End Sub


Private Sub DatumVerwerking_Change()

    MIM_GLOBAL_DATE = Format(DatumVerwerking.Value, "dd/mm/yyyy")
    Mim.DatumVerwerking.Value = DatumVerwerking.Value
        
End Sub


Private Sub Form_Activate()

    If Format(Now, "dd/mm/yyyy") <> Format(DatumVerwerking.Value, "dd/mm/yyyy") Then
        DatumVerwerking.Value = Format(Now, "dd/mm/yyyy")
        DatumVerwerking_Change
    End If
    
End Sub

Private Sub PeriodeBoekjaar_Click()
Dim A As String
Dim aa As String * 4
Dim FlTemp As Integer

A = PeriodeBoekjaar.Text
PERIOD_FROMTO = Mid(A, 7, 4) + Mid(A, 4, 2) + Left(A, 2) + Right(A, 4) + Mid(A, 17, 2) + Mid(A, 14, 2)

FlTemp = FreeFile
Open PROGRAM_LOCATION + "9999.OCT" For Random As FlTemp Len = 4
aa = Str$(ACTIVE_BOOKYEAR)
Put FlTemp, 1, aa
aa = Boekjaar.Text
Put FlTemp, 2, aa
aa = Str$(PeriodeBoekjaar.ListIndex + 1)
Put FlTemp, 3, aa
Close FlTemp

Caption = "(" + Boekjaar.Text + ") (" + Left(PeriodeBoekjaar.Text, 10) + ") BoekPeriode"
End Sub

Private Sub Verkleinen_Click()

BJPERDAT.WindowState = 1
Close FL99
FL99 = 0

End Sub

