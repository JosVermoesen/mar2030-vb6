VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFiguren 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "marFiguurDefinitie"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   8820
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox lbFiguren 
      Height          =   450
      Left            =   6120
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   4620
      Width           =   2655
   End
   Begin VB.CommandButton cmdBewaren 
      Caption         =   "Bewaren"
      Height          =   735
      Left            =   4800
      Picture         =   "Figuren.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Hudige figuurdefinities opslaan in bedrijfsinhoudsopgave"
      Top             =   4380
      Width           =   1215
   End
   Begin VB.TextBox txtHeight 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Text            =   "1200"
      ToolTipText     =   "Afdrukhoogte in Twips"
      Top             =   4680
      Width           =   975
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Openen"
      Height          =   750
      Left            =   60
      Picture         =   "Figuren.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Figuur inladen"
      Top             =   4320
      Width           =   960
   End
   Begin VB.TextBox txtWidth 
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Text            =   "3600"
      ToolTipText     =   "Afdrukbreedte in Twips"
      Top             =   4680
      Width           =   735
   End
   Begin VB.CommandButton cmdFiguur 
      Caption         =   "PrinterTest"
      Height          =   750
      Left            =   3540
      Picture         =   "Figuren.frx":0CD4
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Afdruktest hoogte x breedte via DocumentPrinter"
      Top             =   4380
      Width           =   1215
   End
   Begin MSComCtl2.UpDown udWidth 
      Height          =   375
      Left            =   3195
      TabIndex        =   1
      Top             =   4680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   2200
      BuddyControl    =   "txtWidth"
      BuddyDispid     =   196613
      OrigLeft        =   3840
      OrigTop         =   540
      OrigRight       =   4080
      OrigBottom      =   1335
      Increment       =   100
      Max             =   10000
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.UpDown udHeight 
      Height          =   375
      Left            =   2100
      TabIndex        =   5
      Top             =   4680
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   661
      _Version        =   393216
      Value           =   3500
      BuddyControl    =   "txtHeight"
      BuddyDispid     =   196611
      OrigLeft        =   3840
      OrigTop         =   540
      OrigRight       =   4080
      OrigBottom      =   1335
      Increment       =   100
      Max             =   18000
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Bestaande marFiguurdefinities"
      Height          =   195
      Index           =   2
      Left            =   6180
      TabIndex        =   10
      Top             =   4380
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Breedte"
      Height          =   195
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Hoogte"
      Height          =   195
      Index           =   0
      Left            =   1140
      TabIndex        =   6
      Top             =   4440
      Width           =   615
   End
   Begin VB.Image imgFiguur 
      BorderStyle     =   1  'Fixed Single
      Height          =   1200
      Left            =   60
      Stretch         =   -1  'True
      Top             =   60
      Width           =   3600
   End
End
Attribute VB_Name = "frmFiguren"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FigBestandsnaam As String


Sub PrintAnywhere(Src As Object, Dest As Object)
    If Dest Is Printer Then
        Dest.PaintPicture Src.Picture, 0, 0, Src.Width, Src.Height
        Printer.EndDoc
    Else
        Dest.PaintPicture Src.Picture, Dest.Width / 2, Dest.Height / 2
    End If
   
End Sub


Private Sub cmdBewaren_Click()
Dim ktrlString As String

If FigBestandsnaam = "" Then Beep: Exit Sub

Msg = "Figuur op te slaan in bedrijfsinhoudsopgave "
Msg = Msg & LOCATION_COMPANYDATA
Msg = Msg & " en voorkeurinstellingen "
Msg = Msg & txtHeight & " / " & txtWidth & vbCr & vbCr
Msg = Msg & "GEEF NAAM VOOR DEFINITIE"
ktrlString = InputBox(Msg, "FiguurDefinitie bewaren", "Figuur")
If ktrlString <> "" Then
    If ScrMaakTekstBestand(FigBestandsnaam & vbTab & txtHeight & vbTab & txtWidth, LOCATION_COMPANYDATA & ktrlString & ".mfd") Then
    Else
        MsgBox "Fout tijdens wegschrijven Figuurdefinitie"
    End If
End If
LijstVernieuwen

End Sub

Private Sub cmdFiguur_Click()

    Screen.MousePointer = vbHourglass
    Set Printer = Printers(dokumentPrinterNr)
    On Error Resume Next
    Printer.PaperBin = LaadTekst(App.Title, "dokumentPRINTER")
    If Printer.Orientation = vbPRORLandscape Then
        Printer.Orientation = vbPRORPortrait
        DoEvents
    End If
    PrintAnywhere imgFiguur, Printer
    Screen.MousePointer = vbNormal

End Sub

Function VernieuwTekening()

    imgFiguur.Height = txtHeight
    imgFiguur.Width = txtWidth
    
End Function

Private Sub Form_Load()

    LijstVernieuwen
    VernieuwTekening
    If lbFiguren.ListCount Then
        lbFiguren.ListIndex = 0
    End If
    
End Sub

Private Sub imgFiguur_DblClick()

imgFiguur.Stretch = Not imgFiguur.Stretch
If imgFiguur.Stretch = False Then
    txtHeight.Text = imgFiguur.Height
    txtWidth.Text = imgFiguur.Width
End If

End Sub

Private Sub lbFiguren_Click()
    
    On Local Error Resume Next
    Msg = ""
    If ScrLeesTekstBestand(Msg, LOCATION_COMPANYDATA & lbFiguren.Text & ".mfd") Then
        FigBestandsnaam = Left(Msg, InStr(Msg, vbTab) - 1)
        imgFiguur.Picture = LoadPicture(FigBestandsnaam)
        frmFiguren.Caption = "marFiguurDefinitie: " & FigBestandsnaam
        Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
        txtHeight = Left(Msg, InStr(Msg, vbTab) - 1)
        Msg = Mid(Msg, InStr(Msg, vbTab) + 1)
        txtWidth = Msg
        VernieuwTekening
    End If
    
End Sub

Private Sub lbFiguren_KeyDown(KeyCode As Integer, Shift As Integer)

    If lbFiguren.ListIndex <> -1 And KeyCode = 46 Then
        Msg = LOCATION_COMPANYDATA & lbFiguren & ".mfd" & vbCr & vbCr
        Msg = Msg & "Bent U zeker"
        KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, "Verwijderen !")
        If KtrlBox = vbYes Then
            If Dir(LOCATION_COMPANYDATA & lbFiguren & ".mfd") <> "" Then
                Kill LOCATION_COMPANYDATA & lbFiguren & ".mfd"
                lbFiguren.RemoveItem lbFiguren.ListIndex
                If lbFiguren.ListCount <> 0 Then
                    lbFiguren.ListIndex = 0
                End If
            End If
        End If
    End If
    
    
End Sub

Private Sub txtHeight_LostFocus()

    VernieuwTekening
    
End Sub


Private Sub txtWidth_LostFocus()
    
    VernieuwTekening
    
End Sub

Private Sub udHeight_Change()

    VernieuwTekening

End Sub

Private Sub udWidth_Change()

    VernieuwTekening

End Sub

Private Sub cmdOpen_Click()

    On Error GoTo CancelError
    Mim.Teken.filename = ""
    Mim.Teken.CancelError = True
    Mim.Teken.Filter = _
        "Alle Figuurbestanden|*.bmp;*.dib;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur" & _
        "|JPEG bestanden (*.jpg)|*.jpg" & _
        "|GIF bestanden (*.gif)|*.gif" & _
        "|BITMAP bestanden (*.bmp;*.dib)|*.bmp;*.dib" & _
        "|META bestanden (*.wmf, *.emf)|*.wmf;*.emf" & _
        "|CURSOR bestanden (*.ico, *.cur)|*.ico;*.cur"
                
    Mim.Teken.ShowOpen
    imgFiguur.Picture = LoadPicture(Mim.Teken.filename)
    If Err Then
    Else
        FigBestandsnaam = Mim.Teken.filename
        frmFiguren.Caption = "marFiguurDefinitie: " & FigBestandsnaam
    End If
    
CancelError:
Screen.MousePointer = vbNormal

End Sub

Function LijstVernieuwen()
Dim Bestanden As String

lbFiguren.Clear
Bestanden = Dir(LOCATION_COMPANYDATA & "*.mfd")
Do While Bestanden <> ""
    lbFiguren.AddItem Left(Bestanden, Len(Bestanden) - 4)
    Bestanden = Dir
Loop

End Function
