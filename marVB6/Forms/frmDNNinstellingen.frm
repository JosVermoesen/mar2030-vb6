VERSION 5.00
Begin VB.Form frmDNNinstellingen 
   Caption         =   "DNN & Postvak Instellingen"
   ClientHeight    =   3135
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9450
   LinkTopic       =   "Form3"
   ScaleHeight     =   3135
   ScaleWidth      =   9450
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton ButtonCodaIOMap 
      Height          =   315
      Left            =   8880
      Picture         =   "frmDNNinstellingen.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Bureaublad openen"
      Top             =   600
      Width           =   435
   End
   Begin VB.TextBox txtCodaIOMap 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   17
      Text            =   "Text1"
      ToolTipText     =   "Voorbeeld: http://rvherdersem.no-ip.biz/rvdnn"
      Top             =   600
      Width           =   6735
   End
   Begin VB.CommandButton ButtonToggle 
      Caption         =   "Toggle Bewerken"
      Height          =   375
      Left            =   6720
      TabIndex        =   15
      Top             =   2520
      Width           =   1515
   End
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
      Left            =   2400
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2520
      Width           =   2295
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
      Left            =   240
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2520
      Width           =   1995
   End
   Begin VB.CommandButton ButtonCloudArchive 
      Height          =   315
      Left            =   8880
      Picture         =   "frmDNNinstellingen.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      ToolTipText     =   "Cloud Archief Locatie openen"
      Top             =   2040
      Width           =   435
   End
   Begin VB.CommandButton ButtonCloudMario 
      Height          =   315
      Left            =   8880
      Picture         =   "frmDNNinstellingen.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Mario Locatie openen"
      Top             =   1560
      Width           =   435
   End
   Begin VB.CommandButton ButtonCloudMarnt 
      Height          =   315
      Left            =   8880
      Picture         =   "frmDNNinstellingen.frx":0306
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Marnt Cloud Locatie openen"
      Top             =   1080
      Width           =   435
   End
   Begin VB.TextBox txtSkyDriveMap 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Text            =   "Text1"
      ToolTipText     =   "Voorbeeld: c:\rvdnn\portals\0\documenten\postvak in"
      Top             =   1080
      Width           =   6735
   End
   Begin VB.CommandButton cmdSluiten 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   8400
      TabIndex        =   7
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtPDFpostvak 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Text            =   "Text1"
      ToolTipText     =   "Voorbeeld: c:\rvdnn\portals\0\documenten\postvak in"
      Top             =   1560
      Width           =   6735
   End
   Begin VB.CommandButton CmdBewaar 
      Caption         =   "&Bewaren en sluiten"
      Height          =   375
      Left            =   4920
      TabIndex        =   6
      Top             =   2520
      Width           =   1635
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
      Top             =   2040
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "CODA I/O MAP"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "CLOUD MARNT"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "CLOUD MARIO"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
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
      Top             =   2040
      Width           =   1815
   End
End
Attribute VB_Name = "frmDNNinstellingen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z

Dim toggleEdit As Boolean

Sub ToggleProperties(toggleSet As Boolean)

    ButtonDefaultResetForMapMarnt.Visible = toggleSet
    ButtonDefaultResetForOneDrive.Visible = toggleSet
    CmdBewaar.Visible = toggleSet
    
    txtURLwww.Enabled = toggleSet
    txtSkyDriveMap.Enabled = toggleSet
    txtPDFpostvak.Enabled = toggleSet
    txtURLlokaal.Enabled = toggleSet
    
End Sub

Private Sub ButtonCloudArchive_Click()
    
    If ShellExecuteWithFallback(Me.txtURLlokaal.text) = False Then
        MsgBox "Kon " & Me.txtURLlokaal.text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
    End If
    
End Sub

Private Sub ButtonCloudMario_Click()

    If ShellExecuteWithFallback(Me.txtPDFpostvak.text) = False Then
        MsgBox "Kon " & Me.txtPDFpostvak.text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
    End If
    
End Sub

Private Sub ButtonCloudMarnt_Click()
    
    If ShellExecuteWithFallback(Me.txtSkyDriveMap.text) = False Then
        MsgBox "Kon " & Me.txtSkyDriveMap.text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
    End If
    
End Sub

Private Sub ButtonCodaIOMap_Click()

    If ShellExecuteWithFallback(Me.txtCodaIOMap.text) = False Then
        MsgBox "Kon " & Me.txtCodaIOMap.text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
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
        Me.txtSkyDriveMap.text = marNTLocatie
        Me.txtPDFpostvak.text = marNTLocatie + "\manueel"
        Me.txtURLlokaal.text = marNTLocatie + "\archief"
        
        On Error Resume Next
        Err = 0
        fs.CreateFolder (Me.txtPDFpostvak.text)
        
        If Err Then
            MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & Me.txtPDFpostvak.text, vbInformation
        End If
        On Error Resume Next
        Err = 0
        
        fs.CreateFolder (Me.txtURLlokaal.text)
        If Err Then
            MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & Me.txtURLlokaal.text, vbInformation
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
            Me.txtSkyDriveMap.text = systemPersonalDocs & "\marNT"
            Me.txtPDFpostvak.text = systemPersonalDocs & "\marNT\manueel"
            Me.txtURLlokaal.text = systemPersonalDocs & "\marNT\archief"
            
            On Error Resume Next
            Err = 0
            fs.CreateFolder (Me.txtSkyDriveMap.text)
            If Err Then
                MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & Me.txtSkyDriveMap.text, vbInformation
            End If
        
            On Error Resume Next
            Err = 0
            fs.CreateFolder (Me.txtPDFpostvak.text)
            If Err Then
                MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & Me.txtPDFpostvak.text, vbInformation
            End If
            
            On Error Resume Next
            Err = 0
            fs.CreateFolder (Me.txtURLlokaal.text)
            If Err Then
                MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & Me.txtURLlokaal.text, vbInformation
            End If
                        
            Call CmdBewaar_Click
        End If
    End If
    
End Sub

Private Sub ButtonToggle_Click()

    toggleEdit = Not toggleEdit
    ToggleProperties (toggleEdit)
    If toggleEdit = True Then
        MsgBox "Wees bedachtzaam bij het wijzigen van deze belangrijke instellingen voor MarIntegraal en MarSync", vbExclamation
    End If
    
End Sub



Private Sub CmdBewaar_Click()

    BeWaarTekst "dnnInstellingen", "Archief", Me.txtURLlokaal.text 'archief cloud
    BeWaarTekst "dnnInstellingen", "URLwww", Me.txtURLwww.text
    BeWaarTekst "dnnInstellingen", "Mario", Me.txtPDFpostvak.text 'mario cloud
    BeWaarTekst "dnnInstellingen", "Cloud", Me.txtSkyDriveMap.text 'marnt cloud
    Unload Me

End Sub

Private Sub cmdSluiten_Click()

    Unload Me
    
End Sub


Private Sub Form_Load()
    
    toggleEdit = False
    ToggleProperties (toggleEdit)
        
    If LOCATION_COMPANYDATA = "" Then
    Else
        Me.ButtonDefaultResetForOneDrive.Enabled = True
        Me.ButtonDefaultResetForMapMarnt.Enabled = True
    End If
    
    If LaadTekst("dnnInstellingen", "CodaIOMap") = "" Then
        BeWaarTekst "dnnInstellingen", "CodaIOMap", LOCATION_DESKTOP
        Me.txtCodaIOMap.text = LaadTekst("dnnInstellingen", "CodaIOMap")
    End If

    If LaadTekst("dnnInstellingen", "Cloud") = "" Then
        Dim bedrijfsLoc As String
        bedrijfsLoc = LaadTekst(App.Title, "Bedrijfsinhoudsopgave")
        MsgBox "Nieuwe PC of nog geen instellingen voor Cloud.  Wijzig de volgende standaardwaarden a.u.b. voor uw bedrijf (zie aanbevelingen in onze voorbeeld nota!) of vraag onze gratis bijstand om dit in uw plaats in orde te brengen.", vbInformation
        Me.txtURLlokaal.text = bedrijfsLoc & "\cloud\archief" ' "http://localhost/rvDNN"
        Me.txtURLwww.text = "https://mijndomein.be"
        Me.txtSkyDriveMap.text = bedrijfsLoc & "\cloud" '"C:\Users\NaamVanGebruiker\SkyDrive"
        Me.txtPDFpostvak.text = bedrijfsLoc & "\cloud\mario" '"c:\dotnetnuke\rvDNN\portals\0\documenten\postvak"
    Else
        Me.txtURLlokaal.text = LaadTekst("dnnInstellingen", "Archief")
        Me.txtURLwww.text = LaadTekst("dnnInstellingen", "URLwww")
        Me.txtSkyDriveMap.text = LaadTekst("dnnInstellingen", "Cloud")
        Me.txtPDFpostvak.text = LaadTekst("dnnInstellingen", "Mario")
        Me.txtCodaIOMap.text = LaadTekst("dnnInstellingen", "CodaIOMap")
    End If

End Sub

