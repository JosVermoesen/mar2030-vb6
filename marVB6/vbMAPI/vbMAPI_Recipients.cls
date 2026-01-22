VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMAPI_Recipients"
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

Property Get Count() As Long: End Property
Attribute Count.VB_UserMemId = -101
Property Get Item(ByVal IndexOrName As Variant) As vbMAPI_Recipient: End Property
Attribute Item.VB_UserMemId = 0
Function Add(Optional ByVal NameOrObject As Variant, Optional ByVal RecipientType As EnumRecipientType = Recipient_TO) As vbMAPI_Recipient: End Function
Attribute Add.VB_UserMemId = -102
Function AddEx(ByVal Name As String, Optional ByVal Address As String, Optional ByVal AddressType As String, Optional ByVal RecipientType As EnumRecipientType = Recipient_TO) As vbMAPI_Recipient: End Function
Attribute AddEx.VB_UserMemId = -103
Property Get Session() As vbMAPI_Session: End Property
Attribute Session.VB_UserMemId = -104
Sub ShowAddressBook(Optional ByVal Title As String, Optional ByVal OneAddress As Boolean, Optional ByVal ForceResolution As Boolean = True, Optional ByVal NumRecipGroups As Long = &HFFFFFFFF, Optional ByVal ToCaption As String, Optional ByVal CcCaption As String, Optional ByVal BccCaption As String, Optional ByVal ParentWindowHandle As Long): End Sub
Attribute ShowAddressBook.VB_UserMemId = -105
Function ResolveAll(Optional ByVal ShowDialog As Boolean = True, Optional ByVal ParentWindowHandle As Long) As Boolean: End Function
Attribute ResolveAll.VB_UserMemId = -106
Sub Clear(): End Sub
Attribute Clear.VB_UserMemId = -107
