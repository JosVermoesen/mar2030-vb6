VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMAPI_EncodedBody"
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

Enum EnumEncodedBodyFormat
    BodyFormat_Unknown = &H0
    BodyFormat_PlainText = &H1
    BodyFormat_HTML = &H2
    BodyFormat_RTF = &H3
End Enum

Enum EnumBodySource
    BodySourceFrom_AUTO = &HFFFFFFFF
    BodySourceFrom_PlainText = &H1
    BodySourceFrom_HTML = &H2
    BodySourceFrom_RTF = &H3
End Enum

Property Get BodyFormat() As EnumEncodedBodyFormat: End Property
Attribute BodyFormat.VB_UserMemId = -101
Property Get BodySource() As EnumBodySource: End Property
Attribute BodySource.VB_UserMemId = -102
Property Get Value() As String: End Property
Attribute Value.VB_UserMemId = 0
Property Let Value(ByVal Value As String): End Property
Attribute Value.VB_UserMemId = 0
Property Get CodePage() As Long: End Property
Attribute CodePage.VB_UserMemId = -103
