VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "msmapi32.ocx"
Begin VB.Form frmMailTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Outlook Mail Test"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   9405
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cbGetFile 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   19
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton cbMapiSendMail 
      Caption         =   "Testmail versturen met MAPI"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3600
      TabIndex        =   18
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9045
      _ExtentX        =   15954
      _ExtentY        =   8916
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Test Verzenden"
      TabPicture(0)   =   "frmOutlookMailTest.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbFile"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbAttachment"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "tbBody"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tbSubject"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cbSendMail"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Test Ontvangen"
      TabPicture(1)   =   "frmOutlookMailTest.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(2)=   "LabelAantalBerichten"
      Tab(1).Control(3)=   "rtbMail"
      Tab(1).Control(4)=   "TextBoxRetours"
      Tab(1).Control(5)=   "TextBoxCheckForMail"
      Tab(1).Control(6)=   "cbMailBerichten"
      Tab(1).Control(7)=   "ButtonFetch2"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton ButtonFetch2 
         Caption         =   "Fetch2"
         Height          =   495
         Left            =   -74760
         TabIndex        =   15
         Top             =   480
         Width           =   1215
      End
      Begin VB.ComboBox cbMailBerichten 
         Height          =   315
         ItemData        =   "frmOutlookMailTest.frx":0038
         Left            =   -74760
         List            =   "frmOutlookMailTest.frx":003A
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1320
         Width           =   8535
      End
      Begin VB.TextBox TextBoxCheckForMail 
         Height          =   285
         Left            =   -68040
         TabIndex        =   11
         Text            =   "peppol@vsoft.be"
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TextBoxRetours 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   -68040
         TabIndex        =   10
         Text            =   "0"
         Top             =   750
         Width           =   1815
      End
      Begin VB.CommandButton cbSendMail 
         Caption         =   "Testmail versturen met OUTLOOK Object"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   4260
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.TextBox tbSubject 
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Text            =   "MailTest"
         Top             =   540
         Width           =   8775
      End
      Begin VB.TextBox tbBody 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Text            =   "frmOutlookMailTest.frx":003C
         Top             =   1260
         Width           =   8775
      End
      Begin RichTextLib.RichTextBox rtbMail 
         Height          =   3015
         Left            =   -74760
         TabIndex        =   17
         Top             =   1800
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   5318
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmOutlookMailTest.frx":0085
      End
      Begin VB.Label LabelAantalBerichten 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Totaal aantal berichten: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -73440
         TabIndex        =   16
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Te Controleren Mailbox"
         Height          =   255
         Left            =   -70200
         TabIndex        =   13
         Top             =   510
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "Ontvangen Retour"
         Height          =   255
         Left            =   -69840
         TabIndex        =   12
         Top             =   750
         Width           =   1695
      End
      Begin VB.Label lbAttachment 
         Caption         =   "Bijlage(n) (*.pdf, *.xml]"
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
         Left            =   240
         TabIndex        =   8
         Top             =   2820
         Width           =   2175
      End
      Begin VB.Label lbFile 
         BorderStyle     =   1  'Fixed Single
         Height          =   1095
         Left            =   120
         TabIndex        =   7
         Top             =   3060
         Width           =   7935
      End
      Begin VB.Label Label2 
         Caption         =   "Inhoud:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1020
         Width           =   975
      End
   End
   Begin VB.CommandButton cbCheckMapi 
      Caption         =   "Check voor MAPI (aanbevolen)"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   2535
   End
   Begin VB.CommandButton cbCheckOutlook 
      Caption         =   "Check voor OUTLOOK Object"
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   5280
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSMAPI.MAPIMessages MPIBericht 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MPISessie 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Label Label1 
      Caption         =   "Onderwerp:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frmMailTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
DefInt A-Z
Option Explicit

    Dim objOutlook As Outlook.Application
    Dim objMail As Outlook.MailItem
    Dim objAttach As Outlook.Attachment
    
    Dim mapiMailBijlage As String
    
    Dim aTTachFileName() As String
    Dim aTTachFullName() As String

    Dim checkForMapiCaption As String
    Dim checkFor0utlookObjectCaption As String
    

Private Sub ButtonFetch2_Click()

    On Error Resume Next
    
    Dim testRunning As Boolean
    testRunning = outlookIsRunning()
    If outlookIsRunning = False Then
        Exit Sub
    End If
    
ProbeerNogEens:
    DoEvents
    Me.MPIBericht.Fetch
    If Err = 32053 Then
        GoSub mailLogon
        GoTo ProbeerNogEens
    Else
        Me.refreshmailMAPIMessages
        Me.cbMailBerichten.ListIndex = 0
    End If
    DoEvents
    Exit Sub
    
mailLogon:
    Err = 0
    Me.MPISessie.DownLoadMail = True
    Me.MPISessie.SignOn
    If Err Then
        MsgBox "Mail niet beschikbaar" & vbCrLf & Error, vbInformation
        Me.cbMailBerichten.Enabled = False
    Else
        Me.MPIBericht.SessionID = Me.MPISessie.SessionID
    End If
Return

End Sub

Function refreshmailMAPIMessages() As Boolean

    Me.MPIBericht.FetchMsgType = "IPM"
    Me.MPIBericht.Fetch

    Me.LabelAantalBerichten.Caption = "Totaal aantal berichten:" & vbCrLf & Me.MPIBericht.MsgCount
    Me.cbMailBerichten.Clear
    Me.rtbMail.TextRTF = ""
    
    Dim aantalRB As Integer
    Dim cbMailTekstLijn As String
    
    aantalRB = 0
    Me.cbMailBerichten.Clear
    If Me.MPIBericht.MsgCount > 0 Then
        On Error Resume Next
        For COUNT_TO = 0 To Me.MPIBericht.MsgCount - 1
            Me.MPIBericht.MsgIndex = COUNT_TO
            'MsgBox "Stop"
            ' TODO
            ' eerst kijken of het mailadres in orde is
            If InStr(Me.MPIBericht.RecipAddress, Me.TextBoxCheckForMail.Text) Then
                cbMailTekstLijn = "ok PEPPOL-mail:=" & Me.MPIBericht.RecipAddress  'mailMAPIMessages.MsgSubject
                'dan controleren of er attachment is met TB2
                If Me.MPIBericht.AttachmentCount <> 0 Then
                    'MsgBox "Stop bij: er zijn attachments"
                
                     aantalRB = aantalRB + 1
                    ' ? mailMAPIMessages.AttachmentPathName
                    ' C:\Users\Jos\AppData\Local\Temp\BR298247.txt
                    Me.rtbMail.LoadFile Me.MPIBericht.AttachmentName
                    cbMailTekstLijn = cbMailTekstLijn & " Bijlage Gevonden : " & Me.MPIBericht.AttachmentPathName
               Else
                    'msgbox "Stop bij: er zijn geen attachments in dit bericht
                    cbMailTekstLijn = cbMailTekstLijn & " Geen Bijlage Gevonden/Ongeldig: " & Me.MPIBericht.MsgSubject
                End If
            Else
                cbMailTekstLijn = "nok PEPPOL-mail:=" & Me.MPIBericht.RecipAddress & " : " & Me.MPIBericht.MsgSubject
            End If
    
            If Err Then Exit For
            Me.cbMailBerichten.AddItem cbMailTekstLijn ', Me.cbMailBerichten.ListCount - 1
        Next
        'MsgBox "stop"

        Me.TextBoxRetours.Text = Str(aantalRB)
        On Error GoTo 0
        If Me.cbMailBerichten.ListCount Then
            Me.cbMailBerichten.ListIndex = 0
            On Error Resume Next
            Me.cbMailBerichten.SetFocus
        End If
    End If

End Function


Private Sub cbCheckMapi_Click()
    
    Dim result As Boolean
    
    result = SignOnMapi(True)
    If result Then
        MsgBox "MAPI Control en Sessie met succes opgestart. IDkode :" + Format(Me.MPISessie.SessionID), vbInformation
        Me.cbCheckMapi.Caption = checkForMapiCaption + ": OK!"
        Me.cbMapiSendMail.Visible = True
        Me.cbCheckMapi.Enabled = False
        MsgBox "U werkt met Outlook op PC (classic)?" & vbCrLf & vbCrLf & "Vermijd problemen en open Outlook eerst (mag offline staan)", vbExclamation
    End If
    
End Sub

Function SignOnMapi(downloadFlag As Boolean) As Boolean

    SignOnMapi = False

    Err = 0
    On Local Error Resume Next
    If downloadFlag Then
        Me.MPISessie.DownLoadMail = True
    End If
    Me.MPISessie.SignOn
    If Err Then
        MsgBox Err.Description, vbCritical
        MsgBox "Om mail rechtstreeks vanuit marIntegraal te verzenden dient U te beschikken over een email programma.", vbInformation
        Exit Function
    Else
        Err = 0
        On Local Error Resume Next
        Me.MPIBericht.SessionID = Me.MPISessie.SessionID
        If Err Then
            MsgBox Err.Description, vbCritical
            Exit Function
        Else
            SignOnMapi = True
        End If
    End If
    
End Function

Private Sub cbCheckOutlook_Click()
    
    Err = 0
    On Local Error Resume Next
    Set objOutlook = New Outlook.Application
    If Err Then
        MsgBox Err.Description, vbCritical
        MsgBox "Om mail rechtstreeks vanuit marIntegraal te verzenden dient U te beschikken over een email programma.", vbInformation
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    
    Dim OutlookVersion As String
    
    Err = 0
    On Local Error Resume Next
    OutlookVersion = objOutlook.Version
    
    Screen.MousePointer = Default
    If Err Then
        MsgBox Err.Description
        Set objOutlook = Nothing
        Exit Sub
    Else
        MsgBox "Outlook version is: " & OutlookVersion
        cbCheckOutlook.Caption = checkFor0utlookObjectCaption & ": OK!"
        cbSendMail.Visible = True
        cbCheckOutlook.Enabled = False
    End If
    
    Dim testRunning As Boolean
    testRunning = outlookIsRunning()
    If outlookIsRunning = False Then
        
    
    End If
        
    objOutlook.Quit
    Set objOutlook = Nothing
        
End Sub

Function outlookIsRunning() As Boolean
    
    outlookIsRunning = False
    
    Dim oOL As New Outlook.Application
    On Error Resume Next
    Set oOL = GetObject(, "Outlook.Application")
    On Error GoTo 0

    If oOL Is Nothing Then
        Msg = "Outlook zelf moet draaien voor verzending d.m.v. het Outlook Object" & vbCrLf & vbCrLf
        Msg = Msg & "Start Outlook eerst en zet eventueel eerst in offline modus"
        MsgBox Msg, vbInformation
    Else
        outlookIsRunning = True
    End If
    Set oOL = Nothing

End Function


Private Sub cbMailBerichten_Click()

    Me.MPIBericht.MsgIndex = Me.cbMailBerichten.ListIndex
    If InStr(Me.MPIBericht.RecipAddress, Me.TextBoxCheckForMail.Text) Then
        If Me.MPIBericht.AttachmentCount <> 0 Then
            Me.rtbMail.LoadFile Me.MPIBericht.AttachmentPathName
            'Me.cbMailManueel.Enabled = True
        Else
            Me.rtbMail.Text = "Er is geen TB2 bijlage?"
            'Me.cbMailManueel.Enabled = False
        End If
    Else
        Me.rtbMail.Text = Me.MPIBericht.MsgNoteText
        'Me.cbMailManueel.Enabled = False
    End If
    Me.cbMailBerichten.SetFocus
    
End Sub

Private Sub cbMapiSendMail_Click()
       
    If Me.MPISessie.SessionID <> 0 Then
        Me.MousePointer = vbHourglass
        With Me.MPIBericht
            'Compose new message
            .Compose
            
            'Address message
            .RecipDisplayName = "Vsoft Administratieve Software"
            .RecipAddress = "SMTP:" & "peppol@vsoft.be"
        
            'Resolve recipient name
            .AddressResolveUI = True '=dialogbox, false = error genereren
                
            'Create the message
            .MsgSubject = Me.tbSubject.Text & " - " & Now & " met MAPI"
            .MsgNoteText = Me.tbBody.Text & "MAPI"
            '.AttachmentPathName = mapiMailBijlage
        End With
        ' Attach files in a loop
        Dim counter As Integer
        For counter = 0 To UBound(aTTachFileName) - 1
            MPIBericht.AttachmentIndex = counter
            MPIBericht.AttachmentPosition = counter
            MPIBericht.AttachmentName = aTTachFullName(counter)
            MPIBericht.AttachmentPathName = aTTachFileName(counter)
        Next counter
        
        'Send the message
        On Error Resume Next
        Me.MPIBericht.Send True
        Me.MousePointer = vbNormal
        If Err = 32001 Then
            'MsgBox "U kiest de mail niet te verzenden.", vbInformation
        ElseIf Err Then
            MsgBox Error
        End If
        Me.MPISessie.SignOff
    Else
        MsgBox "Eerst sessie opstarten a.u.b.", vbInformation
    End If
        
End Sub

Private Sub cbSendMail_Click()
    
    Err = 0
    On Local Error Resume Next
    
    Set objOutlook = New Outlook.Application
    If Err Then
        MsgBox "Outlook is niet beschiktbaar. Om mail rechtstreeks vanuit marIntegraal met Outlook te verzenden deze eerst installeren a.u.b.", vbInformation
        Exit Sub
    End If
    
    Set objMail = Outlook.CreateItem(olMailItem)
  
    Dim strToAddress As String
    Dim strSubject As String
    Dim strBody As String
  
    strToAddress = "peppol@vsoft.be"
    strSubject = Me.tbSubject.Text & " - " & Now & " OUTLOOK Object"
    strBody = Me.tbBody.Text & "Outlook Object"

    With objMail
        .To = strToAddress
        .Subject = strSubject
        .BodyFormat = olFormatPlain
        .Body = strBody
    End With
  
    ' Add attachment
    For COUNT_TO = 0 To UBound(aTTachFullName) - 1
        objMail.Attachments.Add aTTachFullName(COUNT_TO)
    Next COUNT_TO
    'Set objAttach = objMail.Attachments.Add(Me.lbFile.Caption)
    'objAttach.DisplayName = "MyAttachment"

    'MsgBox "outlook security will now complain as I try to resolve your email addresses against your address book"
    objMail.Recipients.ResolveAll

    objMail.Display
        
    If Err Then
        MsgBox Err & " " & Err.Description
    End If
    Set objMail = Nothing
    Set objAttach = Nothing
    objOutlook.Quit
    Set objOutlook = Nothing
    Exit Sub

SomeError:
    If Err = 0 Then
        'MsgBox "Mail staat klaar in Outlook", vbInformation
    Else
        MsgBox Err & " " & Err.Description, vbCritical
    End If
    Set objMail = Nothing
    Set objAttach = Nothing
    objOutlook.Quit
    Set objOutlook = Nothing

End Sub

Private Sub cbGetFile_Click()
               
    Dim allFiles As String
    
    On Local Error Resume Next
    With Mim.Teken
        ' Set InitialDirectory
        .InitDir = "C:\"
        ' Set CancelError is True
        .CancelError = True
        ' Set flags
        .FLAGS = cdlOFNAllowMultiselect + cdlOFNExplorer
        ' Set filters
        .Filter = "Pdf Bestanden (*.pdf)|*.pdf|Xml Bestanden" & "(*.xml)|*.xml"
        ' Specify default filter
        .FilterIndex = 1
    End With
    
    Mim.Teken.ShowOpen
    If Err Then
        Mim.Teken.filename = " "
    Else
        allFiles = Mim.Teken.filename
        'Parse and use the filelist
        Dim parts() As String
        Dim folderPath As String
        Dim i As Integer
        
        Dim addToOneStringUrls As String
        Dim addToOneStringNames As String
                        
        ' Split on nulls
        parts = Split(allFiles, vbNullChar)

        If UBound(parts) = 0 Then
            ' Single file selected: parts(0) already includes full path
            addToOneStringUrls = parts(0)
            addToOneStringNames = "Bijlage"
        Else
            folderPath = parts(0)
            For i = 1 To UBound(parts)
                ' Combine path + filename
                addToOneStringUrls = addToOneStringUrls & folderPath & "\" & parts(i) & ";"
                addToOneStringNames = addToOneStringNames & parts(i) & ";"
            Next i
        End If
        aTTachFullName = Split(addToOneStringUrls, ";")
        aTTachFileName = Split(addToOneStringNames, ";")
        Me.lbFile.Caption = ""
        Msg = ""
        
        Dim T As Integer
        T = 0
        While T < UBound(aTTachFullName) - 1
            Msg = Msg & aTTachFullName(T) & vbCrLf
            T = T + 1
        Wend
        Msg = Msg & aTTachFullName(T)
        Me.lbFile.Caption = Msg
        
        If Me.MPISessie.SessionID <> 0 Then
            Me.cbMapiSendMail.Enabled = True
        End If
        
        If outlookIsRunning = True Then
            Me.cbSendMail.Enabled = True
        End If
    End If
    
End Sub


Private Sub Form_Load()

    checkForMapiCaption = "Check voor MAPI (aanbevolen)"
    checkFor0utlookObjectCaption = "Check voor OUTLOOK Object"
    
    Msg = "Outlook Object checkup & MAPI checkup" & vbCrLf & vbCrLf
    Msg = Msg & "Testmail zal bij OUTLOOK Object klaar staan " & vbCrLf & "binnen de map 'Drafts' of 'Concepten'" & vbCrLf & vbCrLf
    Msg = Msg & "Testmail zal bij MAPI in een venster ter verzending " & vbCrLf & "voorgesteld worden." & vbCrLf & vbCrLf
    Msg = Msg & "Beide voor controle vooraleer te verzenden." & vbCrLf & vbCrLf
    Msg = Msg & "MAPI verzendwijze blijft meest aanbevolen werkwijze."
    'MsgBox Msg, vbInformation, "Outlook Testen"

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Me.MPISessie.SessionID = 0 Then
    Else
        On Local Error Resume Next
        Me.MPISessie.SignOff
        If Err Then
            MsgBox "Fout tijdens afsluiten van een sessie...  Postvak In (InBox) werd voortijdig gesloten."
        End If
    End If
    
End Sub
