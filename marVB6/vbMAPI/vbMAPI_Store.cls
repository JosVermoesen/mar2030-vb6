VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMAPI_Store"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' *****************************************************************************
' *  Virtual-COM library: vbMAPI  http://www.everythingaccess.com/vbmapi.htm  *
' *  v1.16.3.0 DEMO              built Nov  8 2023                            *
' *---------------------------------------------------------------------------*
' *  This is a demonstration version of the library for TESTING purposes only.*
' *****************************************************************************

Option Explicit

Enum EnumDefaultFolderType
    FolderType_DeletedItems = &H3
    FolderType_Outbox = &H4
    FolderType_SentMail = &H5
    FolderType_Inbox = &H6
    FolderType_Calendar = &H9
    FolderType_Contacts = &HA
    FolderType_Journal = &HB
    FolderType_Notes = &HC
    FolderType_Tasks = &HD
    FolderType_Drafts = &H10
    FolderType_PublicFoldersAllPublicFolders = &H12
    FolderType_Conflicts = &H13
    FolderType_SyncIssues = &H14
    FolderType_LocalFailures = &H15
    FolderType_ServerFailures = &H16
    FolderType_Junk = &H17
End Enum

Event OnNewMail(ByVal EntryID As String)
Property Get Properties() As vbMAPI_Properties: End Property
Attribute Properties.VB_UserMemId = -101
Property Get NamedProperties(ByVal PropertySetGUID As String, ByVal ID As Variant, ByVal DataType As EnumMAPIPropDataTypes) As vbMAPI_Property: End Property
Attribute NamedProperties.VB_UserMemId = -102
Property Get Name() As String: End Property
Attribute Name.VB_UserMemId = 0
Property Get MdbProvider() As String: End Property
Attribute MdbProvider.VB_UserMemId = -103
Property Get EntryID() As String: End Property
Attribute EntryID.VB_UserMemId = -104
Property Get SearchKey() As String: End Property
Attribute SearchKey.VB_UserMemId = -105
Function GetDefaultFolderEntryID(ByVal DefaultFolder As EnumDefaultFolderType) As String: End Function
Attribute GetDefaultFolderEntryID.VB_UserMemId = -106
Function GetDefaultFolder(ByVal DefaultFolder As EnumDefaultFolderType) As vbMAPI_Folder: End Function
Attribute GetDefaultFolder.VB_UserMemId = -107
Property Get Session() As vbMAPI_Session: End Property
Attribute Session.VB_UserMemId = -108
Function GetItemFromID(ByVal EntryID As String) As vbMAPI_MailItem: End Function
Attribute GetItemFromID.VB_UserMemId = -109
Function GetFolderFromID(ByVal EntryID As String) As vbMAPI_Folder: End Function
Attribute GetFolderFromID.VB_UserMemId = -110
Function CompareEntryIDs(ByVal EntryID1 As String, ByVal EntryID2 As String) As Boolean: End Function
Attribute CompareEntryIDs.VB_UserMemId = -111
Property Get IPMRootFolder() As vbMAPI_Folder: End Property
Attribute IPMRootFolder.VB_UserMemId = -112
Property Get RootFolder() As vbMAPI_Folder: End Property
Attribute RootFolder.VB_UserMemId = -113
Sub AbortSubmit(ByVal EntryID As String): End Sub
Attribute AbortSubmit.VB_UserMemId = -114
Property Get MAPIOBJECT() As Variant: End Property
Attribute MAPIOBJECT.VB_UserMemId = -115
