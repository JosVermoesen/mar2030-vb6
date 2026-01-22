VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4950
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   7995
   LinkTopic       =   "Form2"
   ScaleHeight     =   4950
   ScaleWidth      =   7995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cbGrid 
      Caption         =   "Fetch Contacten"
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf 
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   5741
      _Version        =   393216
      Rows            =   5
      Cols            =   5
      FixedCols       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   5
   End
   Begin VB.CommandButton cbContactPersonen 
      Caption         =   "Test ContactPersonen"
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cbAgenda 
      Caption         =   "Test Agenda"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox tbText 
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oApp As Object
Dim oApp2 As Object

Dim ol As Object

Private Sub cbAgenda_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo OutlookAgendaError
    Set oApp2 = CreateObject("Outlook.Application")
    
    
    ' Get NameSpace and Logon.
    Dim oNS As Outlook.NameSpace
    Set oNS = oApp2.GetNamespace("mapi")
    oNS.Logon ("Outlook")
    'oNS.Logon("Outlook", Missing.Value, False, True) ' TODO:

    ' Create and Open a new contact.
    Dim olItem As Outlook.ContactItem
    Set olItem = oApp2.CreateItem(olContactItem)
  

    ' Setup Contact information...
    With olItem
        .FirstName = "Joske"
        .LastName = "Vermeulen"
        '.FullName = "Joske Vermeulen"
        .Birthday = "25/02/1961"
        .CompanyName = "Vsoft"
        .HomeTelephoneNumber = "+3293629685"
        .Email1Address = "josmobile@rv.be"
        .JobTitle = "Ontwikkelaar"
        .HomeAddress = "Molenstraat 8" & vbCrLf & "Balegem, 9860"
        .Categories = "marIntegraal"
    End With

    ' Save Contact...
    olItem.Save
    MsgBox olItem.ConversationIndex
    
    ' Create a new appointment.
    Dim olAppt As Outlook.AppointmentItem
    Set olAppt = oApp2.CreateItem(olAppointmentItem)

    ' Set start time for 2-minutes from now...
    olAppt.Start = Now
        
    ' Setup other appointment information...
    With olAppt
        .Duration = 60
        .Subject = "Vergadering om plannen te bespreken..."
        .Body = "Vergadering met " & olItem.FullName & " om de plannen te bespreken."
        .Location = "Thuis Kantoor"
        .ReminderMinutesBeforeStart = 1
        .ReminderSet = True
    End With

    ' Save Appointment...
    olAppt.Save
    
    ' Send a message to your new contact.
    Dim olMail As Outlook.MailItem
    Set olMail = oApp2.CreateItem(olMailItem)
    ' Fill out & send message...
    olMail.To = olItem.Email1Address
    olMail.Subject = "Over onze vergadering..."
    olMail.Body = _
         "Beste " & olItem.FirstName & ", " & vbCr & vbCr & vbTab & _
         "Ik zie U binnen 2 minuutjes op onze vergadering!" & vbCr & vbCr & _
         "Ps: Ik heb U toegevoegd aan mijn lijst contactpersonen."
    olMail.Send
    
    ' Clean up...
    MsgBox "Alle taken volbracht...", vbInformation
    
    oNS.Logoff
    'oNS = Nothing
        
OutlookAgendaError:
    Screen.MousePointer = vbNormal
    

End Sub

Private Sub cbContactPersonen_Click()

    Screen.MousePointer = vbHourglass
    On Error GoTo OutLookErrorCTRL
    'Set oApp = New Outlook.Application
        
    Set oApp = CreateObject("Outlook.Application")
      
    ' Get NameSpace and Logon.
    Dim oNS As Outlook.NameSpace
    Set oNS = oApp.GetNamespace("mapi")
    oNS.Logon ("Outlook")
    'oNS.Logon("Outlook", Missing.Value, False, True) ' TODO:
    
    ' Get the first contact from the Contacts folder.
    Dim cContacts As Outlook.MAPIFolder
    Set cContacts = oNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts)
    
    Dim oItems As Outlook.Items
    Set oItems = cContacts.Items
    
    Dim oCt As Outlook.ContactItem
    Set oCt = oItems.GetFirst()
           
    Label1.Caption = oItems.Count
    
    ' Display some common properties.
    tbText.Text = (oCt.FullName) & vbCrLf
    tbText.Text = tbText.Text & (oCt.Birthday) & vbCrLf
    tbText.Text = tbText.Text & (oCt.CompanyName) & vbCrLf
    tbText.Text = tbText.Text & (oCt.Department) & vbCrLf
    tbText.Text = tbText.Text & (oCt.Body) & vbCrLf
    tbText.Text = tbText.Text & (oCt.FileAs) & vbCrLf
    tbText.Text = tbText.Text & (oCt.Email1Address) & vbCrLf
    tbText.Text = tbText.Text & (oCt.Subject) & vbCrLf
    tbText.Text = tbText.Text & (oCt.JobTitle) & vbCrLf
    tbText.Text = tbText.Text & (oCt.ConversationIndex) & vbCrLf
    tbText.Text = tbText.Text & (oCt.EntryID) & vbCrLf
       
    Debug.Print oCt.BusinessFaxNumber
    
    ' Save Contact...
    'oCt.Save
    MsgBox tbText

OutLookErrorCTRL:
    Screen.MousePointer = vbNormal
    If Err Then MsgBox Error
    

End Sub

Private Sub cbGrid_Click()

    Dim strTest As String
    
    Me.msf.Cols = 7
        
    strTest = "test1" + Chr(9) + "test2" + Chr(9) + "Test1b" + vbTab + "Test2b"
    Me.msf.TextMatrix(0, 0) = "CompanyName"
    Me.msf.TextMatrix(0, 1) = "LastNameAndFirstName"
    Me.msf.TextMatrix(0, 2) = "Email1Address"
    Me.msf.TextMatrix(0, 3) = "BusinessTelephoneNumber"
    Me.msf.TextMatrix(0, 4) = "HomeTelephoneNumber"
    Me.msf.TextMatrix(0, 5) = "BusinessFaxNumber"
    Me.msf.TextMatrix(0, 6) = "EntryID"
    

    
End Sub


Private Sub Command1_Click()
   'The Outlook object library must be referenced
   'Dim ol As Outlook.Application
   'ol onder general als object!
      
   Dim olns As NameSpace
   Dim objFolder As Object
   Dim AllContacts As Object
   Dim Item As Object
   Dim i As Integer
   
   'If there are more than 500 contacts, change the following line:
   Dim EntryID(1000) As String
   Dim FullName(1000) As String
   Dim CompanyName(1000) As String
   Dim Phone(1000) As String
   Dim Fax(1000) As String
   Dim Mobile(1000) As String
   Dim Email(1000) As String
   
   Dim StoreID As String
   Dim strFind As String
   
   ' Set the application object
   Set ol = New Outlook.Application
   
   ' Set the namespace object
   Set olns = ol.GetNamespace("MAPI")
   ' Set the default Contacts folder
   Set objFolder = olns.GetDefaultFolder(olFolderContacts)
   ' Get the StoreID, which is a property of the folder
   StoreID = objFolder.StoreID
   ' Set objAllContacts = the collection of all contacts
   Set AllContacts = objFolder.Items
   i = 0
   ' Loop to get all of the EntryIDs for the contacts
   For Each Item In AllContacts
      i = i + 1
      ' The EntryID is a property of the item
      EntryID(i) = Item.EntryID
      FullName(i) = Item.FullName
      'If InStr(UCase(FullName(i)), "BIEKEN") Then
      '  bGet FlKlant, 0, "010018"
      '  If Ktrl = 0 Then
      '      rsMAR(FlKlant)("v262") = EntryID(i)
      '      rsMAR(FlKlant).Update
      '  End If
      'End If
      
      CompanyName(i) = Item.CompanyName
      Phone(i) = Item.BusinessTelephoneNumber
      Mobile(i) = Item.MobileTelephoneNumber
      Fax(i) = Item.BusinessFaxNumber
      Email(i) = Item.Email1Address
   Next
   ' Randomly choose the 2nd Contact to retrieve
   ' In a larger solution, this might be the index from a listbox
   i = 57
   ' Both the StoreID and EntryID must be used to retrieve the item
   Set Item = olns.GetItemFromID(EntryID(i), StoreID)
   'Item.Display
   'Stop

End Sub
