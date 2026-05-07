VERSION 5.00
Begin VB.Form frmDNNinstellingen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DNN & Postvak Instellingen"
   ClientHeight    =   4950
   ClientLeft      =   105
   ClientTop       =   405
   ClientWidth     =   8790
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   8790
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame FrameCoda 
      Caption         =   "Coda XML en XDA In/Uit"
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   3960
      Width           =   8535
      Begin VB.TextBox TextBoxCodaIOMap 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   21
         Text            =   "Text1"
         ToolTipText     =   "Voorbeeld: http://rvherdersem.no-ip.biz/rvdnn"
         Top             =   360
         Width           =   6735
      End
      Begin VB.CommandButton ButtonCodaIOMap 
         Height          =   315
         Left            =   7920
         Picture         =   "frmDNNinstellingen.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Coda Map locatie openen"
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "MAP"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame FrameCompanyData 
      Caption         =   "Locatie Bedrijfmappen"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   3000
      Width           =   8535
      Begin VB.TextBox TextBoxMarntDataMap 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   17
         Text            =   "Text1"
         ToolTipText     =   "Voorbeeld: http://rvherdersem.no-ip.biz/rvdnn"
         Top             =   360
         Width           =   6735
      End
      Begin VB.CommandButton ButtonMarntDataMap 
         Height          =   315
         Left            =   7920
         Picture         =   "frmDNNinstellingen.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   16
         TabStop         =   0   'False
         ToolTipText     =   "Marnt DATA Locatie openen"
         Top             =   360
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "DATA"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame FrameCloud 
      Caption         =   "Cloud"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   8535
      Begin VB.TextBox TextBoxCloudArchive 
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Text            =   "Text1"
         ToolTipText     =   "voorbeeld: http://localhost/rvdnn"
         Top             =   1320
         Width           =   6735
      End
      Begin VB.TextBox TextBoxCloudMario 
         Height          =   285
         Left            =   1080
         TabIndex        =   10
         Text            =   "Text1"
         ToolTipText     =   "Voorbeeld: c:\rvdnn\portals\0\documenten\postvak in"
         Top             =   840
         Width           =   6735
      End
      Begin VB.TextBox TextBoxCloudMarnt 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Text            =   "Text1"
         ToolTipText     =   "Voorbeeld: c:\rvdnn\portals\0\documenten\postvak in"
         Top             =   360
         Width           =   6735
      End
      Begin VB.CommandButton ButtonCloudMarnt 
         Height          =   315
         Left            =   7920
         Picture         =   "frmDNNinstellingen.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Marnt Cloud Locatie openen"
         Top             =   360
         Width           =   435
      End
      Begin VB.CommandButton ButtonCloudMario 
         Height          =   315
         Left            =   7920
         Picture         =   "frmDNNinstellingen.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Mario Locatie openen"
         Top             =   840
         Width           =   435
      End
      Begin VB.CommandButton ButtonCloudArchive 
         Height          =   315
         Left            =   7920
         Picture         =   "frmDNNinstellingen.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Cloud Archief Locatie openen"
         Top             =   1320
         Width           =   435
      End
      Begin VB.Label lblURLlokaal 
         Caption         =   "ARCHIEF"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "MARIO"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "MARNT"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.CommandButton ButtonToggle 
      Caption         =   "Toggle Bewerken"
      Height          =   375
      Left            =   5880
      TabIndex        =   4
      Top             =   2040
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
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1455
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
      Height          =   735
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1275
   End
   Begin VB.CommandButton ButtonClose 
      Cancel          =   -1  'True
      Caption         =   "Sluiten"
      Height          =   375
      Left            =   7560
      TabIndex        =   1
      Top             =   2040
      Width           =   975
   End
   Begin VB.CommandButton ButtonSave 
      Caption         =   "&Bewaren en sluiten"
      Height          =   735
      Left            =   3120
      TabIndex        =   0
      Top             =   2040
      Width           =   1275
   End
End
Attribute VB_Name = "frmDNNinstellingen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

Dim toggleEdit As Boolean

Sub ToggleProperties(toggleSet As Boolean)

    ButtonDefaultResetForMapMarnt.Visible = toggleSet
    ButtonDefaultResetForOneDrive.Visible = toggleSet
    ButtonSave.Visible = toggleSet
        
    TextBoxCloudMarnt.Enabled = toggleSet
    TextBoxCloudMario.Enabled = toggleSet
    TextBoxCloudArchive.Enabled = toggleSet
    
End Sub

Private Sub ButtonCloudArchive_Click()
    
    If ShellExecuteWithFallback(TextBoxCloudArchive.text) = False Then
        MsgBox "Kon " & TextBoxCloudArchive.text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
    End If
    
End Sub

Private Sub ButtonCloudMario_Click()

    If ShellExecuteWithFallback(TextBoxCloudMario.text) = False Then
        MsgBox "Kon " & TextBoxCloudMario.text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
    End If
    
End Sub

Private Sub ButtonCloudMarnt_Click()
    
    If ShellExecuteWithFallback(TextBoxCloudMarnt.text) = False Then
        MsgBox "Kon " & TextBoxCloudMarnt.text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
    End If
    
End Sub

Private Sub ButtonCodaIOMap_Click()

    If ShellExecuteWithFallback(TextBoxCodaIOMap.text) = False Then
        MsgBox "Kon " & TextBoxCodaIOMap.text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
    End If

End Sub

Private Sub ButtonMarntDataMap_Click()
    
    If ShellExecuteWithFallback(TextBoxMarntDataMap.text) = False Then
        MsgBox "Kon " & TextBoxMarntDataMap.text & " niet openen. Raadpleeg ShellHelper.log voor details.", vbExclamation
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
    
    Msg = "Akkoord voor:" & vbCrLf
    Msg = Msg & "CLOUD   MARNT: " & marNTLocatie & " (dus dezelfde hoofdmap)" & vbCrLf
    Msg = Msg & "CLOUD   MARIO: " & marNTLocatie & "\manueel" & vbCrLf
    Msg = Msg & "CLOUD ARCHIEF: " & marNTLocatie & "\archief" & vbCrLf
    
    KtrlBox = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
    If KtrlBox = vbYes Then
        TextBoxCloudMarnt.text = marNTLocatie
        TextBoxCloudMario.text = marNTLocatie + "\manueel"
        TextBoxCloudArchive.text = marNTLocatie + "\archief"
        
        On Error Resume Next
        Err = 0
        fs.CreateFolder (TextBoxCloudMario.text)
        
        If Err Then
            MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & TextBoxCloudMario.text, vbInformation
        End If
        On Error Resume Next
        Err = 0
        
        fs.CreateFolder (TextBoxCloudArchive.text)
        If Err Then
            MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & TextBoxCloudArchive.text, vbInformation
        End If
        Call ButtonSave_Click
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
            TextBoxCloudMarnt.text = systemPersonalDocs & "\marNT"
            TextBoxCloudMario.text = systemPersonalDocs & "\marNT\manueel"
            TextBoxCloudArchive.text = systemPersonalDocs & "\marNT\archief"
            
            On Error Resume Next
            Err = 0
            fs.CreateFolder (TextBoxCloudMarnt.text)
            If Err Then
                MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & TextBoxCloudMarnt.text, vbInformation
            End If
        
            On Error Resume Next
            Err = 0
            fs.CreateFolder (TextBoxCloudMario.text)
            If Err Then
                MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & TextBoxCloudMario.text, vbInformation
            End If
            
            On Error Resume Next
            Err = 0
            fs.CreateFolder (TextBoxCloudArchive.text)
            If Err Then
                MsgBox "Map bestaat reeds" & vbCrLf & vbCrLf & TextBoxCloudArchive.text, vbInformation
            End If
                        
            Call ButtonSave_Click
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

Private Sub ButtonSave_Click()

    BeWaarTekst "dnnInstellingen", "Archief", TextBoxCloudArchive.text 'archief cloud
    
    BeWaarTekst "dnnInstellingen", "Mario", TextBoxCloudMario.text 'mario cloud
    BeWaarTekst "dnnInstellingen", "Cloud", TextBoxCloudMarnt.text 'marnt cloud
    Unload Me

End Sub

Private Sub ButtonClose_Click()

    Unload Me
    
End Sub


Private Sub Form_Load()

    Dim strDataLocatie As String

    strDataLocatie = LaadTekst("BedrijfOpenen", "DataDefault")
    If strDataLocatie = "server" Then
        TextBoxMarntDataMap.text = LaadTekst(App.Title, "ServerBedrijfsinhoudsopgave")  'Server anders
    Else
        TextBoxMarntDataMap.text = LaadTekst(App.Title, "Bedrijfsinhoudsopgave" & "2025") 'Lokaal is standaard
    End If
    
    toggleEdit = False
    ToggleProperties (toggleEdit)
        
    If LOCATION_COMPANYDATA = "" Then
    Else
        ButtonDefaultResetForOneDrive.Enabled = True
        ButtonDefaultResetForMapMarnt.Enabled = True
    End If
    
    If LaadTekst("dnnInstellingen", "CodaIOMap") = "" Then
        BeWaarTekst "dnnInstellingen", "CodaIOMap", LOCATION_DESKTOP
        TextBoxCodaIOMap.text = LaadTekst("dnnInstellingen", "CodaIOMap")
    End If

    If LaadTekst("dnnInstellingen", "Cloud") = "" Then
        Dim bedrijfsLoc As String
        bedrijfsLoc = LaadTekst(App.Title, "Bedrijfsinhoudsopgave")
        MsgBox "Nieuwe PC of nog geen instellingen voor Cloud.  Wijzig de volgende standaardwaarden a.u.b. voor uw bedrijf (zie aanbevelingen in onze voorbeeld nota!) of vraag onze gratis bijstand om dit in uw plaats in orde te brengen.", vbInformation
        TextBoxCloudArchive.text = bedrijfsLoc & "\cloud\archief" ' "http://localhost/rvDNN"
        
        TextBoxCloudMarnt.text = bedrijfsLoc & "\cloud" '"C:\Users\NaamVanGebruiker\SkyDrive"
        TextBoxCloudMario.text = bedrijfsLoc & "\cloud\mario" '"c:\dotnetnuke\rvDNN\portals\0\documenten\postvak"
    Else
        TextBoxCloudArchive.text = LaadTekst("dnnInstellingen", "Archief")
        
        TextBoxCloudMarnt.text = LaadTekst("dnnInstellingen", "Cloud")
        TextBoxCloudMario.text = LaadTekst("dnnInstellingen", "Mario")
        TextBoxCodaIOMap.text = LaadTekst("dnnInstellingen", "CodaIOMap")
    End If

End Sub


