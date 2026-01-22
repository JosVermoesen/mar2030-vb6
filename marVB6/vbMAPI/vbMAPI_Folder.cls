VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMAPI_Folder"
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

Enum EnumMessageItemType
    ItemType_MailItem = &H0
    ItemType_AppointmentItem = &H1
    ItemType_ContactItem = &H2
    ItemType_TaskItem = &H3
    ItemType_JournalItem = &H4
    ItemType_NoteItem = &H5
    ItemType_PostItem = &H6
    ItemType_DistributionListItem = &H7
End Enum

Enum EnumFolderKind
    FolderKind_Root = &H0
    FolderKind_Generic = &H1
    FolderKind_Search = &H2
End Enum

Property Get Properties() As vbMAPI_Properties: End Property
Attribute Properties.VB_UserMemId = -101
Property Get NamedProperties(ByVal PropertySetGUID As String, ByVal ID As Variant, ByVal DataType As EnumMAPIPropDataTypes) As vbMAPI_Property: End Property
Attribute NamedProperties.VB_UserMemId = -102
Property Get Items() As vbMAPI_FolderItems: End Property
Attribute Items.VB_UserMemId = -103
Property Get Folders() As vbMAPI_Folders: End Property
Attribute Folders.VB_UserMemId = -104
Property Get DefaultItemType() As EnumMessageItemType: End Property
Attribute DefaultItemType.VB_UserMemId = -105
Property Get DefaultMessageClass() As String: End Property
Attribute DefaultMessageClass.VB_UserMemId = -106
Property Get ContainerClass() As String: End Property
Attribute ContainerClass.VB_UserMemId = -107
Property Let ContainerClass(ByVal RetVal As String): End Property
Attribute ContainerClass.VB_UserMemId = -107
Property Get Name() As String: End Property
Attribute Name.VB_UserMemId = 0
Property Let Name(ByVal RetVal As String): End Property
Attribute Name.VB_UserMemId = 0
Property Get Description() As String: End Property
Attribute Description.VB_UserMemId = -108
Property Let Description(ByVal RetVal As String): End Property
Attribute Description.VB_UserMemId = -108
Property Get EntryID() As String: End Property
Attribute EntryID.VB_UserMemId = -109
Property Get ParentEntryID() As String: End Property
Attribute ParentEntryID.VB_UserMemId = -110
Property Get StoreEntryID() As String: End Property
Attribute StoreEntryID.VB_UserMemId = -111
Property Get FolderKind() As EnumFolderKind: End Property
Attribute FolderKind.VB_UserMemId = -112
Property Get UnReadItemCount() As Long: End Property
Attribute UnReadItemCount.VB_UserMemId = -113
Property Get ContentCount() As Long: End Property
Attribute ContentCount.VB_UserMemId = -114
Property Get SearchKey() As String: End Property
Attribute SearchKey.VB_UserMemId = -115
Property Get HasSubfolders() As Boolean: End Property
Attribute HasSubfolders.VB_UserMemId = -116
Property Get Parent() As vbMAPI_Folder: End Property
Attribute Parent.VB_UserMemId = -117
Property Get Store() As vbMAPI_Store: End Property
Attribute Store.VB_UserMemId = -118
Property Get Session() As vbMAPI_Session: End Property
Attribute Session.VB_UserMemId = -119
Property Get MAPIOBJECT() As Variant: End Property
Attribute MAPIOBJECT.VB_UserMemId = -120
Sub Delete(Optional ByVal DeleteSubFolders As Boolean = True, Optional ByVal DeleteItems As Boolean = True): End Sub
Attribute Delete.VB_UserMemId = -121
