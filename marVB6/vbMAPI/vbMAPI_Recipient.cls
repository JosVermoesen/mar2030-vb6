VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMAPI_Recipient"
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

Property Get Name() As String: End Property
Attribute Name.VB_UserMemId = 0
Property Let Name(ByVal RetVal As String): End Property
Attribute Name.VB_UserMemId = 0
Property Get EmailAddress() As String: End Property
Attribute EmailAddress.VB_UserMemId = -101
Property Let EmailAddress(ByVal RetVal As String): End Property
Attribute EmailAddress.VB_UserMemId = -101
Property Get EntryID() As String: End Property
Attribute EntryID.VB_UserMemId = -102
Property Let EntryID(ByVal RetVal As String): End Property
Attribute EntryID.VB_UserMemId = -102
Property Get RecipientType() As EnumRecipientType: End Property
Attribute RecipientType.VB_UserMemId = -103
Property Let RecipientType(ByVal RetVal As EnumRecipientType): End Property
Attribute RecipientType.VB_UserMemId = -103
Property Get AddressType() As String: End Property
Attribute AddressType.VB_UserMemId = -104
Property Let AddressType(ByVal RetVal As String): End Property
Attribute AddressType.VB_UserMemId = -104
Property Get Flags() As Long: End Property
Attribute Flags.VB_UserMemId = -105
Property Let Flags(ByVal RetVal As Long): End Property
Attribute Flags.VB_UserMemId = -105
Property Get ObjectType() As Long: End Property
Attribute ObjectType.VB_UserMemId = -106
Property Let ObjectType(ByVal RetVal As Long): End Property
Attribute ObjectType.VB_UserMemId = -106
Property Get SearchKey() As String: End Property
Attribute SearchKey.VB_UserMemId = -107
Property Get DisplayType() As Long: End Property
Attribute DisplayType.VB_UserMemId = -108
Property Let DisplayType(ByVal RetVal As Long): End Property
Attribute DisplayType.VB_UserMemId = -108
Property Get SendRichInfo() As Boolean: End Property
Attribute SendRichInfo.VB_UserMemId = -109
Property Let SendRichInfo(ByVal Value As Boolean): End Property
Attribute SendRichInfo.VB_UserMemId = -109
Sub Delete(): End Sub
Attribute Delete.VB_UserMemId = -110
Property Get IsResolved() As Boolean: End Property
Attribute IsResolved.VB_UserMemId = -111
Function Resolve(Optional ByVal ShowDialog As Boolean = True, Optional ByVal ParentWndHandle As Long) As Boolean: End Function
Attribute Resolve.VB_UserMemId = -112
Sub ResolveOneOff(): End Sub
Attribute ResolveOneOff.VB_UserMemId = -113
Property Get Session() As vbMAPI_Session: End Property
Attribute Session.VB_UserMemId = -114
Property Get AddressEntry() As Object: End Property
Attribute AddressEntry.VB_UserMemId = -115
