VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMAPI_Attachment"
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

Enum EnumAttachmentMethod
    AttachmentMethod_ByValue = &H1
    AttachmentMethod_ByReference = &H4
    AttachmentMethod_Embedded = &H5
    AttachmentMethod_OLE = &H6
End Enum

Property Get Properties() As vbMAPI_Properties: End Property
Attribute Properties.VB_UserMemId = -101
Property Get NamedProperties(ByVal PropertySetGUID As String, ByVal ID As Variant, ByVal DataType As EnumMAPIPropDataTypes) As vbMAPI_Property: End Property
Attribute NamedProperties.VB_UserMemId = -102
Property Get AttachmentNumber() As Long: End Property
Attribute AttachmentNumber.VB_UserMemId = -103
Property Get DisplayName() As String: End Property
Attribute DisplayName.VB_UserMemId = 0
Property Let DisplayName(ByVal Value As String): End Property
Attribute DisplayName.VB_UserMemId = 0
Property Get FileName() As String: End Property
Attribute FileName.VB_UserMemId = -104
Property Let FileName(ByVal Value As String): End Property
Attribute FileName.VB_UserMemId = -104
Property Get Hidden() As Boolean: End Property
Attribute Hidden.VB_UserMemId = -105
Property Let Hidden(ByVal RetVal As Boolean): End Property
Attribute Hidden.VB_UserMemId = -105
Property Get PathName() As String: End Property
Attribute PathName.VB_UserMemId = -106
Property Get Position() As Long: End Property
Attribute Position.VB_UserMemId = -107
Property Let Position(ByVal RetVal As Long): End Property
Attribute Position.VB_UserMemId = -107
Property Get Size() As Long: End Property
Attribute Size.VB_UserMemId = -108
Property Get Type_() As Long: End Property
Attribute Type_.VB_UserMemId = -109
Property Get CreationTime() As Double: End Property
Attribute CreationTime.VB_UserMemId = -110
Property Get LastModificationTime() As Double: End Property
Attribute LastModificationTime.VB_UserMemId = -111
Property Get ContentID() As String: End Property
Attribute ContentID.VB_UserMemId = -112
Property Let ContentID(ByVal RetVal As String): End Property
Attribute ContentID.VB_UserMemId = -112
Property Get ContentMIMEType() As String: End Property
Attribute ContentMIMEType.VB_UserMemId = -113
Property Let ContentMIMEType(ByVal RetVal As String): End Property
Attribute ContentMIMEType.VB_UserMemId = -113
Property Get ContentLocation() As String: End Property
Attribute ContentLocation.VB_UserMemId = -114
Property Let ContentLocation(ByVal RetVal As String): End Property
Attribute ContentLocation.VB_UserMemId = -114
Property Get IsContactPhoto() As Boolean: End Property
Attribute IsContactPhoto.VB_UserMemId = -115
Property Let IsContactPhoto(ByVal RetVal As Boolean): End Property
Attribute IsContactPhoto.VB_UserMemId = -115
Property Get Flags() As Long: End Property
Attribute Flags.VB_UserMemId = -116
Property Let Flags(ByVal RetVal As Long): End Property
Attribute Flags.VB_UserMemId = -116
Property Get FileSize() As Long: End Property
Attribute FileSize.VB_UserMemId = -117
Sub SaveAsFile(ByVal FilePath As String): End Sub
Attribute SaveAsFile.VB_UserMemId = -118
Property Get DataAsStream() As Variant: End Property
Attribute DataAsStream.VB_UserMemId = -119
Property Get DataAsText() As String: End Property
Attribute DataAsText.VB_UserMemId = -120
Property Let DataAsText(ByVal Data As String): End Property
Attribute DataAsText.VB_UserMemId = -120
Property Get DataAsTextAnsi() As String: End Property
Attribute DataAsTextAnsi.VB_UserMemId = -121
Property Let DataAsTextAnsi(ByVal Data As String): End Property
Attribute DataAsTextAnsi.VB_UserMemId = -121
Property Get DataAsByteArray() As Variant: End Property
Attribute DataAsByteArray.VB_UserMemId = -122
Property Let DataAsByteArray(ByVal DataVar As Variant): End Property
Attribute DataAsByteArray.VB_UserMemId = -122
Property Get Session() As vbMAPI_Session: End Property
Attribute Session.VB_UserMemId = -123
Property Get EmbeddedMessage() As vbMAPI_MailItem: End Property
Attribute EmbeddedMessage.VB_UserMemId = -124
Property Get MAPIOBJECT() As Variant: End Property
Attribute MAPIOBJECT.VB_UserMemId = -125
Sub Delete(): End Sub
Attribute Delete.VB_UserMemId = -126
