VERSION 5.00
Begin VB.Form frmDNNinstellingen 
   Caption         =   "DNN & Postvak Instellingen"
   ClientHeight    =   2580
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9450
   LinkTopic       =   "Form3"
   ScaleHeight     =   2580
   ScaleWidth      =   9450
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ButtonDefaultResetForMapMarnt 
      Caption         =   "AutoDefault MAP MarNT"
      Enabled         =   0   'False
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
      Left            =   2280
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2475
   End
   Begin VB.CommandButton ButtonDefaultResetForOneDrive 
      Caption         =   "AutoDefault OneDrive"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2040
      Width           =   2115
   End
   Begin VB.CommandButton ButtonCloudArchive 
      Height          =   315
      Left            =   8880
      Picture         =   "frmDNNinstellingen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Cloud Archief Locatie openen"
      Top             =   1560
      Width           =   435
   End
   Begin VB.CommandButton ButtonCloudMario 
      Height          =   315
      Left            =   8880
      Picture         =   "frmDNNinstellingen.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Mario Locatie openen"
      Top             =   1080
      Width           =   435
   End
   Begin VB.CommandButton ButtonCloudMarnt 
      Height          =   315
      Left            =   8880
      Picture         =   "frmDNNinstellingen.frx":0294
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Marnt Cloud Locatie openen"
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtSkyDriveMap 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Text            =   "Text1"
      ToolTipText     =   "Voorbeeld: c:\rvdnn\portals\0\documenten\postvak in"
      Top             =   600
      Width           =   6735
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox txtPDFpostvak 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Text            =   "Text1"
      ToolTipText     =   "Voorbeeld: c:\rvdnn\portals\0\documenten\postvak in"
      Top             =   1080
      Width           =   6735
   End
   Begin VB.CommandButton CmdBewaar 
      Caption         =   "Instellingen &bewaren en sluiten"
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   2040
      Width           =   2595
   End
   Begin VB.TextBox txtURLwww 
      Height          =   285
      Left            =   2040
      TabIndex        =   3
      Text            =   "Text1"
      ToolTipText     =   "Voorbeeld: http://rvherdersem.no-ip.biz/rvdnn"
      Top             =   120
      Width           =   6735
   End
   Begin VB.TextBox txtURLlokaal 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Text            =   "Text1"
      ToolTipText     =   "voorbeeld: http://localhost/rvdnn"
      Top             =   1560
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "CLOUD MARNT"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "CLOUD MARIO"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label lblURLwww 
      Caption         =   "URL HOOFDBEDRIJF"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblURLlokaal 
      Caption         =   "CLOUD ARCHIEF"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   1815
   End
End
Attribute VB_Name = "frmDNNinstellingen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Private Sub ButtonCloudArchive_Click()
    
    If ShellExecuteWithFallback(Me.txtURLlokaal.Text) = False Then
        MsgBox "Kon " & Me.txtURLlokaal.Text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
    End If
    
End Sub

Private Sub ButtonCloudMario_Click()

    If ShellExecuteWithFallback(Me.txtPDFpostvak.Text) = False Then
        MsgBox "Kon " & Me.txtPDFpostvak.Text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
    End If
    
End Sub

Private Sub ButtonCloudMarnt_Click()
    
    If ShellExecuteWithFallback(Me.txtSkyDriveMap.Text) = False Then
        MsgBox "Kon " & Me.txtSkyDriveMap.Text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
    End If
    
End Sub

Private Sub ButtonDefaultResetForMapMarnt_Click()

    Dim marNTLocatie As String
    marNTLocatie = LCase(LaadTekst(App.Title, "Bedrijfsinhoudsopgave" & "2025"))
    marNTLocatie = Replace(marNTLocatie, "\data", "")
    
            
    Dim serverMap As String
    serverMap = Trim(LCase(LaadTekst(App.Title, "ServerBedrijfsinhoudsopgave")))
    If serverMap <> "" Then
        Msg = "Voor deze PC bestaat al een serverinhoudsopgave:" & vbCrLf
        Msg = Msg & serverMap & vbCrLf & vbCrLf
        Msg = Msg & "Verwijder indien nodig." & vbCrLf & vbCrLf
        Msg = Msg & "Met serverinstellingen gedraagt deze PC zich als client." & vbCrLf
        Msg = Msg & "Voor marIntegraal draaiende op locatie (client) dient U" & vbCrLf
        Msg = Msg & "de instellingen manueel in te voeren a.d.h.v. uw server-link."
        MsgBox Msg, vbInformation
        Exit Sub
    End If
    
    Msg = "Akkoor voor:" & vbCrLf
    Msg = Msg & "CLOUD   MARNT: " & marNTLocatie & " (dus dezelfde hoofdmap)" & vbCrLf
    Msg = Msg & "CLOUD   MARIO: " & marNTLocatie & "\manueel" & vbCrLf
    Msg = Msg & "CLOUD ARCHIEF: " & marNTLocatie & "\archief" & vbCrLf
    
    KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
    If KtrlBox = vbYes Then
        Me.txtSkyDriveMap.Text = marNTLocatie
        Me.txtPDFpostvak.Text = marNTLocatie + "\manueel"
        Me.txtURLlokaal.Text = marNTLocatie + "\archief"
        
        On Error Resume Next
        Err = 0
        fs.CreateFolder (Me.txtPDFpostvak.Text)
        
        If Err Then
            MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & Me.txtPDFpostvak.Text, vbInformation
        End If
        On Error Resume Next
        Err = 0
        
        fs.CreateFolder (Me.txtURLlokaal.Text)
        If Err Then
            MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & Me.txtURLlokaal.Text, vbInformation
        End If
        Call CmdBewaar_Click
    End If
    
End Sub

Private Sub ButtonDefaultResetForOneDrive_Click()

    Dim marNTLocatie As String
    marNTLocatie = LCase(LaadTekst(App.Title, "Bedrijfsinhoudsopgave" & "2025"))
    marNTLocatie = Replace(marNTLocatie, "\data", "")
            
    Dim serverMap As String
    serverMap = Trim(LCase(LaadTekst(App.Title, "ServerBedrijfsinhoudsopgave")))
    If serverMap <> "" Then
        Msg = "Voor deze PC bestaat al een serverinhoudsopgave:" & vbCrLf
        Msg = Msg & serverMap & vbCrLf & vbCrLf
        Msg = Msg & "Verwijder indien nodig." & vbCrLf & vbCrLf
        Msg = Msg & "Met serverinstellingen gedraagt deze PC zich als client." & vbCrLf
        Msg = Msg & "Voor marIntegraal draaiende op locatie (client) dient U" & vbCrLf
        Msg = Msg & "de instellingen manueel in te voeren a.d.h.v. uw server-link."
        MsgBox Msg, vbInformation
        Exit Sub
    End If
        
    Dim systemPersonalDocs As String
    systemPersonalDocs = LCase(SYSTEM_MYPERSONALDOCUMENTS)
    If InStr(systemPersonalDocs, "onedrive") Then
        Msg = "Dit is een toestel met 'OneDrive' Map ideaal voor automatische" & vbCrLf
        Msg = Msg & "archivering naar de CLOUD." & vbCrLf & vbCrLf
        Msg = Msg & "Akkoord voor:" & vbCrLf
        Msg = Msg & "CLOUD   MARNT: " & systemPersonalDocs & "\marNT" & vbCrLf
        Msg = Msg & "CLOUD   MARIO: " & systemPersonalDocs & "\marNT\manueel" & vbCrLf
        Msg = Msg & "CLOUD ARCHIEF: " & systemPersonalDocs & "\marNT\archief" & vbCrLf
        KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
        If KtrlBox = vbYes Then
            Me.txtSkyDriveMap.Text = systemPersonalDocs & "\marNT"
            Me.txtPDFpostvak.Text = systemPersonalDocs & "\marNT\manueel"
            Me.txtURLlokaal.Text = systemPersonalDocs & "\marNT\archief"
            
            On Error Resume Next
            Err = 0
            fs.CreateFolder (Me.txtSkyDriveMap.Text)
            If Err Then
                MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & Me.txtSkyDriveMap.Text, vbInformation
            End If
        
            On Error Resume Next
            Err = 0
            fs.CreateFolder (Me.txtPDFpostvak.Text)
            If Err Then
                MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & Me.txtPDFpostvak.Text, vbInformation
            End If
            
            On Error Resume Next
            Err = 0
            fs.CreateFolder (Me.txtURLlokaal.Text)
            If Err Then
                MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & Me.txtURLlokaal.Text, vbInformation
            End If
                        
            Call CmdBewaar_Click
        End If
    End If
    
End Sub

Private Sub CmdBewaar_Click()

    BeWaarTekst "dnnInstellingen", "Archief", Me.txtURLlokaal.Text 'archief cloud
    BeWaarTekst "dnnInstellingen", "URLwww", Me.txtURLwww.Text
    BeWaarTekst "dnnInstellingen", "Mario", Me.txtPDFpostvak.Text 'mario cloud
    BeWaarTekst "dnnInstellingen", "Cloud", Me.txtSkyDriveMap.Text 'marnt cloud
    Unload Me

End Sub

Private Sub cmdSluiten_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    If LOCATION_COMPANYDATA = "" Then
    Else
        Me.ButtonDefaultResetForOneDrive.Enabled = True
        Me.ButtonDefaultResetForMapMarnt.Enabled = True
    End If

    If LaadTekst("dnnInstellingen", "Cloud") = "" Then
        Dim bedrijfsLoc As String
        bedrijfsLoc = LaadTekst(App.Title, "Bedrijfsinhoudsopgave")
        MsgBox "Nieuwe PC of nog geen instellingen voor Cloud.  Wijzig de volgende standaardwaarden a.u.b. voor uw bedrijf (zie aanbevelingen in onze voorbeeld nota!) of vraag onze gratis bijstand om dit in uw plaats in orde te brengen.", vbInformation
        Me.txtURLlokaal.Text = bedrijfsLoc & "\cloud\archief" ' "http://localhost/rvDNN"
        Me.txtURLwww.Text = "https://mijndomein.be"
        Me.txtSkyDriveMap.Text = bedrijfsLoc & "\cloud" '"C:\Users\NaamVanGebruiker\SkyDrive"
        Me.txtPDFpostvak.Text = bedrijfsLoc & "\cloud\mario" '"c:\dotnetnuke\rvDNN\portals\0\documenten\postvak"
    Else
        Me.txtURLlokaal.Text = LaadTekst("dnnInstellingen", "Archief")
        Me.txtURLwww.Text = LaadTekst("dnnInstellingen", "URLwww")
        Me.txtSkyDriveMap.Text = LaadTekst("dnnInstellingen", "Cloud")
        Me.txtPDFpostvak.Text = LaadTekst("dnnInstellingen", "Mario")
    End If

End Sub

