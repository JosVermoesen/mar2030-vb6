VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vbMAPI_MailItem"
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

Enum EnumDeleteType
    DeleteType_MoveToDeletedItems = &H0
    DeleteType_SoftDelete = &H1
    DeleteType_HardDelete = &H2
End Enum

Enum EnumFlagIcon
    FlagIcon_NoFlagIcon = &H0
    FlagIcon_PurpleFlagIcon = &H1
    FlagIcon_OrangeFlagIcon = &H2
    FlagIcon_GreenFlagIcon = &H3
    FlagIcon_YellowFlagIcon = &H4
    FlagIcon_BlueFlagIcon = &H5
    FlagIcon_RedFlagIcon = &H6
End Enum

Enum EnumFlagStatus
    FlagStatus_NoFlag = &H0
    FlagStatus_FlagComplete = &H1
    FlagStatus_FlagMarked = &H2
End Enum

Enum EnumImportance
    Importance_Low = &H0
    Importance_Normal = &H1
    Importance_High = &H2
End Enum

Enum EnumSensitivity
    Sensitivity_Normal = &H0
    Sensitivity_Personal = &H1
    Sensitivity_Private = &H2
    Sensitivity_Confidential = &H3
End Enum

Enum EnumRecipientType
    Recipient_TO = &H1
    Recipient_CC = &H2
    Recipient_BCC = &H3
End Enum

Enum EnumMessageFileFormat
    FileFormat_MSG = &H3
    FileFormat_MSGUnicode = &H9
End Enum

Enum EnumImportType
    ImportType_RFC822 = &H400
End Enum

Enum EnumDownloadState
    DownloadState_FullItem = &H1
    DownloadState_HeaderOnly = &H0
End Enum

Enum EnumBusyStatus
    BusyStatus_Free = &H0
    BusyStatus_Tentative = &H1
    BusyStatus_Busy = &H2
    BusyStatus_OutOfOffice = &H3
End Enum

Property Get Properties() As vbMAPI_Properties: End Property
Attribute Properties.VB_UserMemId = -101
Property Get NamedProperties(ByVal PropertySetGUID As String, ByVal ID As Variant, ByVal DataType As EnumMAPIPropDataTypes) As vbMAPI_Property: End Property
Attribute NamedProperties.VB_UserMemId = -102
Property Get Recipients() As vbMAPI_Recipients: End Property
Attribute Recipients.VB_UserMemId = -103
Property Let Recipients(ByVal InRecipients As vbMAPI_Recipients): End Property
Attribute Recipients.VB_UserMemId = -103
Property Get Attachments() As vbMAPI_Attachments: End Property
Attribute Attachments.VB_UserMemId = -104
Sub Save(): End Sub
Attribute Save.VB_UserMemId = -105
Sub Send(): End Sub
Attribute Send.VB_UserMemId = -106
Property Get RTFBody() As vbMAPI_EncodedBody: End Property
Attribute RTFBody.VB_UserMemId = -107
Property Get HTMLBody() As vbMAPI_EncodedBody: End Property
Attribute HTMLBody.VB_UserMemId = -108
Property Get Body() As vbMAPI_EncodedBody: End Property
Attribute Body.VB_UserMemId = -109
Property Get BodyFormat() As EnumEncodedBodyFormat: End Property
Attribute BodyFormat.VB_UserMemId = -110
Property Get SendAccount() As Object: End Property
Attribute SendAccount.VB_UserMemId = -111
Property Get SendAccountStamp() As String: End Property
Attribute SendAccountStamp.VB_UserMemId = -112
Property Let SendAccountStamp(ByVal Value As String): End Property
Attribute SendAccountStamp.VB_UserMemId = -112
Property Get SendAccountName() As String: End Property
Attribute SendAccountName.VB_UserMemId = -113
Property Let SendAccountName(ByVal AccountName As String): End Property
Attribute SendAccountName.VB_UserMemId = -113
Property Get To_() As String: End Property
Attribute To_.VB_UserMemId = -114
Property Let To_(ByVal StringList As String): End Property
Attribute To_.VB_UserMemId = -114
Property Get CC() As String: End Property
Attribute CC.VB_UserMemId = -115
Property Let CC(ByVal StringList As String): End Property
Attribute CC.VB_UserMemId = -115
Property Get BCC() As String: End Property
Attribute BCC.VB_UserMemId = -116
Property Let BCC(ByVal StringList As String): End Property
Attribute BCC.VB_UserMemId = -116
Property Get SaveSentMessageFolderEntryID() As String: End Property
Attribute SaveSentMessageFolderEntryID.VB_UserMemId = -117
Property Let SaveSentMessageFolderEntryID(ByVal RetVal As String): End Property
Attribute SaveSentMessageFolderEntryID.VB_UserMemId = -117
Property Get InternetMessageHeaders() As String: End Property
Attribute InternetMessageHeaders.VB_UserMemId = -118
Property Get AlternateRecipientAllowed() As Boolean: End Property
Attribute AlternateRecipientAllowed.VB_UserMemId = -119
Property Let AlternateRecipientAllowed(ByVal RetVal As Boolean): End Property
Attribute AlternateRecipientAllowed.VB_UserMemId = -119
Property Get AutoForwarded() As Boolean: End Property
Attribute AutoForwarded.VB_UserMemId = -120
Property Let AutoForwarded(ByVal RetVal As Boolean): End Property
Attribute AutoForwarded.VB_UserMemId = -120
Property Get ConversationIndex() As Variant: End Property
Attribute ConversationIndex.VB_UserMemId = -121
Property Let ConversationIndex(ByVal RetVal As Variant): End Property
Attribute ConversationIndex.VB_UserMemId = -121
Property Get ConversationTopic() As String: End Property
Attribute ConversationTopic.VB_UserMemId = -122
Property Let ConversationTopic(ByVal RetVal As String): End Property
Attribute ConversationTopic.VB_UserMemId = -122
Property Get CreationTime() As Date: End Property
Attribute CreationTime.VB_UserMemId = -123
Property Get DeferredDeliveryTime() As Date: End Property
Attribute DeferredDeliveryTime.VB_UserMemId = -124
Property Let DeferredDeliveryTime(ByVal RetVal As Date): End Property
Attribute DeferredDeliveryTime.VB_UserMemId = -124
Property Get DeleteAfterSubmit() As Boolean: End Property
Attribute DeleteAfterSubmit.VB_UserMemId = -125
Property Let DeleteAfterSubmit(ByVal RetVal As Boolean): End Property
Attribute DeleteAfterSubmit.VB_UserMemId = -125
Property Get EntryID() As String: End Property
Attribute EntryID.VB_UserMemId = -126
Property Get ExpiryTime() As Date: End Property
Attribute ExpiryTime.VB_UserMemId = -127
Property Let ExpiryTime(ByVal RetVal As Date): End Property
Attribute ExpiryTime.VB_UserMemId = -127
Property Get FlagIcon() As EnumFlagIcon: End Property
Attribute FlagIcon.VB_UserMemId = -128
Property Let FlagIcon(ByVal RetVal As EnumFlagIcon): End Property
Attribute FlagIcon.VB_UserMemId = -128
Property Get Importance() As EnumImportance: End Property
Attribute Importance.VB_UserMemId = -129
Property Let Importance(ByVal RetVal As EnumImportance): End Property
Attribute Importance.VB_UserMemId = -129
Property Get FlagStatus() As EnumFlagStatus: End Property
Attribute FlagStatus.VB_UserMemId = -130
Property Let FlagStatus(ByVal RetVal As EnumFlagStatus): End Property
Attribute FlagStatus.VB_UserMemId = -130
Property Get InternetCodepage() As Long: End Property
Attribute InternetCodepage.VB_UserMemId = -131
Property Let InternetCodepage(ByVal RetVal As Long): End Property
Attribute InternetCodepage.VB_UserMemId = -131
Property Get LastModificationTime() As Date: End Property
Attribute LastModificationTime.VB_UserMemId = -132
Property Get MessageClass() As String: End Property
Attribute MessageClass.VB_UserMemId = -133
Property Let MessageClass(ByVal RetVal As String): End Property
Attribute MessageClass.VB_UserMemId = -133
Property Get OriginatorDeliveryReportRequested() As Boolean: End Property
Attribute OriginatorDeliveryReportRequested.VB_UserMemId = -134
Property Let OriginatorDeliveryReportRequested(ByVal RetVal As Boolean): End Property
Attribute OriginatorDeliveryReportRequested.VB_UserMemId = -134
Property Get ParentEntryID() As String: End Property
Attribute ParentEntryID.VB_UserMemId = -135
Property Get StoreEntryID() As String: End Property
Attribute StoreEntryID.VB_UserMemId = -136
Property Get ReadReceiptRequested() As Boolean: End Property
Attribute ReadReceiptRequested.VB_UserMemId = -137
Property Let ReadReceiptRequested(ByVal RetVal As Boolean): End Property
Attribute ReadReceiptRequested.VB_UserMemId = -137
Property Get ReceivedByEntryID() As String: End Property
Attribute ReceivedByEntryID.VB_UserMemId = -138
Property Get ReceivedByName() As String: End Property
Attribute ReceivedByName.VB_UserMemId = -139
Property Get ReceivedOnBehalfOfEntryID() As String: End Property
Attribute ReceivedOnBehalfOfEntryID.VB_UserMemId = -140
Property Get ReceivedOnBehalfOfName() As String: End Property
Attribute ReceivedOnBehalfOfName.VB_UserMemId = -141
Property Get ReceivedTime() As Date: End Property
Attribute ReceivedTime.VB_UserMemId = -142
Property Let ReceivedTime(ByVal RetVal As Date): End Property
Attribute ReceivedTime.VB_UserMemId = -142
Property Get RecipientReassignmentProhibited() As Boolean: End Property
Attribute RecipientReassignmentProhibited.VB_UserMemId = -143
Property Let RecipientReassignmentProhibited(ByVal RetVal As Boolean): End Property
Attribute RecipientReassignmentProhibited.VB_UserMemId = -143
Property Get ReplyRecipientNames() As String: End Property
Attribute ReplyRecipientNames.VB_UserMemId = -144
Property Get SenderEmailAddress() As String: End Property
Attribute SenderEmailAddress.VB_UserMemId = -145
Property Let SenderEmailAddress(ByVal RetVal As String): End Property
Attribute SenderEmailAddress.VB_UserMemId = -145
Property Get SenderEmailType() As String: End Property
Attribute SenderEmailType.VB_UserMemId = -146
Property Let SenderEmailType(ByVal RetVal As String): End Property
Attribute SenderEmailType.VB_UserMemId = -146
Property Get SenderName() As String: End Property
Attribute SenderName.VB_UserMemId = -147
Property Let SenderName(ByVal RetVal As String): End Property
Attribute SenderName.VB_UserMemId = -147
Property Get Sensitivity() As EnumSensitivity: End Property
Attribute Sensitivity.VB_UserMemId = -148
Property Let Sensitivity(ByVal RetVal As EnumSensitivity): End Property
Attribute Sensitivity.VB_UserMemId = -148
Property Get SentOn() As Date: End Property
Attribute SentOn.VB_UserMemId = -149
Property Let SentOn(ByVal RetVal As Date): End Property
Attribute SentOn.VB_UserMemId = -149
Property Get SentOnBehalfOfName() As String: End Property
Attribute SentOnBehalfOfName.VB_UserMemId = -150
Property Let SentOnBehalfOfName(ByVal RetVal As String): End Property
Attribute SentOnBehalfOfName.VB_UserMemId = -150
Property Get SentOnBehalfOfEntryID() As String: End Property
Attribute SentOnBehalfOfEntryID.VB_UserMemId = -151
Property Let SentOnBehalfOfEntryID(ByVal RetVal As String): End Property
Attribute SentOnBehalfOfEntryID.VB_UserMemId = -151
Property Get SentOnBehalfOfSearchKey() As String: End Property
Attribute SentOnBehalfOfSearchKey.VB_UserMemId = -152
Property Let SentOnBehalfOfSearchKey(ByVal RetVal As String): End Property
Attribute SentOnBehalfOfSearchKey.VB_UserMemId = -152
Property Get SentOnBehalfOfEmailAddress() As String: End Property
Attribute SentOnBehalfOfEmailAddress.VB_UserMemId = -153
Property Let SentOnBehalfOfEmailAddress(ByVal RetVal As String): End Property
Attribute SentOnBehalfOfEmailAddress.VB_UserMemId = -153
Property Get SentOnBehalfOfAddressType() As String: End Property
Attribute SentOnBehalfOfAddressType.VB_UserMemId = -154
Property Let SentOnBehalfOfAddressType(ByVal RetVal As String): End Property
Attribute SentOnBehalfOfAddressType.VB_UserMemId = -154
Property Get Size() As Long: End Property
Attribute Size.VB_UserMemId = -155
Property Get Subject() As String: End Property
Attribute Subject.VB_UserMemId = 0
Property Let Subject(ByVal RetVal As String): End Property
Attribute Subject.VB_UserMemId = 0
Property Get SenderEntryID() As String: End Property
Attribute SenderEntryID.VB_UserMemId = -156
Property Let SenderEntryID(ByVal RetVal As String): End Property
Attribute SenderEntryID.VB_UserMemId = -156
Property Get SenderSearchKey() As String: End Property
Attribute SenderSearchKey.VB_UserMemId = -157
Property Let SenderSearchKey(ByVal RetVal As String): End Property
Attribute SenderSearchKey.VB_UserMemId = -157
Property Get SearchKey() As String: End Property
Attribute SearchKey.VB_UserMemId = -158
Property Get ReadReceiptEntryID() As String: End Property
Attribute ReadReceiptEntryID.VB_UserMemId = -159
Property Let ReadReceiptEntryID(ByVal RetVal As String): End Property
Attribute ReadReceiptEntryID.VB_UserMemId = -159
Property Get MessageFlags() As Long: End Property
Attribute MessageFlags.VB_UserMemId = -160
Property Let MessageFlags(ByVal RetVal As Long): End Property
Attribute MessageFlags.VB_UserMemId = -160
Property Get IconIndex() As Long: End Property
Attribute IconIndex.VB_UserMemId = -161
Property Let IconIndex(ByVal RetVal As Long): End Property
Attribute IconIndex.VB_UserMemId = -161
Property Get BillingInformation() As String: End Property
Attribute BillingInformation.VB_UserMemId = -162
Property Let BillingInformation(ByVal Value As String): End Property
Attribute BillingInformation.VB_UserMemId = -162
Property Get OutlookInternalVersion() As Long: End Property
Attribute OutlookInternalVersion.VB_UserMemId = -163
Property Let OutlookInternalVersion(ByVal Value As Long): End Property
Attribute OutlookInternalVersion.VB_UserMemId = -163
Property Get HidePaperClip() As Boolean: End Property
Attribute HidePaperClip.VB_UserMemId = -164
Property Let HidePaperClip(ByVal Value As Boolean): End Property
Attribute HidePaperClip.VB_UserMemId = -164
Property Get Mileage() As String: End Property
Attribute Mileage.VB_UserMemId = -165
Property Let Mileage(ByVal Value As String): End Property
Attribute Mileage.VB_UserMemId = -165
Property Get NoAging() As Boolean: End Property
Attribute NoAging.VB_UserMemId = -166
Property Let NoAging(ByVal Value As Boolean): End Property
Attribute NoAging.VB_UserMemId = -166
Property Get OutlookVersion() As String: End Property
Attribute OutlookVersion.VB_UserMemId = -167
Property Let OutlookVersion(ByVal Value As String): End Property
Attribute OutlookVersion.VB_UserMemId = -167
Property Get UseTnef() As Boolean: End Property
Attribute UseTnef.VB_UserMemId = -168
Property Let UseTnef(ByVal Value As Boolean): End Property
Attribute UseTnef.VB_UserMemId = -168
Property Get FlagRequest() As String: End Property
Attribute FlagRequest.VB_UserMemId = -169
Property Let FlagRequest(ByVal Value As String): End Property
Attribute FlagRequest.VB_UserMemId = -169
Property Get FlagDueBy() As Date: End Property
Attribute FlagDueBy.VB_UserMemId = -170
Property Let FlagDueBy(ByVal Value As Date): End Property
Attribute FlagDueBy.VB_UserMemId = -170
Property Get ReminderMinutesBeforeStart() As Long: End Property
Attribute ReminderMinutesBeforeStart.VB_UserMemId = -171
Property Let ReminderMinutesBeforeStart(ByVal Value As Long): End Property
Attribute ReminderMinutesBeforeStart.VB_UserMemId = -171
Property Get ReminderPlaySound() As Boolean: End Property
Attribute ReminderPlaySound.VB_UserMemId = -172
Property Let ReminderPlaySound(ByVal Value As Boolean): End Property
Attribute ReminderPlaySound.VB_UserMemId = -172
Property Get ReminderSet() As Boolean: End Property
Attribute ReminderSet.VB_UserMemId = -173
Property Let ReminderSet(ByVal Value As Boolean): End Property
Attribute ReminderSet.VB_UserMemId = -173
Property Get ReminderSoundFile() As String: End Property
Attribute ReminderSoundFile.VB_UserMemId = -174
Property Let ReminderSoundFile(ByVal Value As String): End Property
Attribute ReminderSoundFile.VB_UserMemId = -174
Property Get ReminderOverrideDefault() As Boolean: End Property
Attribute ReminderOverrideDefault.VB_UserMemId = -175
Property Let ReminderOverrideDefault(ByVal Value As Boolean): End Property
Attribute ReminderOverrideDefault.VB_UserMemId = -175
Property Get Categories() As String: End Property
Attribute Categories.VB_UserMemId = -176
Property Let Categories(ByVal Value As String): End Property
Attribute Categories.VB_UserMemId = -176
Property Get DownloadState() As EnumDownloadState: End Property
Attribute DownloadState.VB_UserMemId = -177
Property Get UnRead() As Boolean: End Property
Attribute UnRead.VB_UserMemId = -178
Property Let UnRead(ByVal Value As Boolean): End Property
Attribute UnRead.VB_UserMemId = -178
Property Get Submitted() As Boolean: End Property
Attribute Submitted.VB_UserMemId = -179
Property Let Submitted(ByVal Value As Boolean): End Property
Attribute Submitted.VB_UserMemId = -179
Property Get Sent() As Boolean: End Property
Attribute Sent.VB_UserMemId = -180
Property Let Sent(ByVal Value As Boolean): End Property
Attribute Sent.VB_UserMemId = -180
Property Get MAPIOBJECT() As Variant: End Property
Attribute MAPIOBJECT.VB_UserMemId = -181
Sub Display(Optional ByVal Modal As Boolean = True, Optional ByVal ParentWindowHandle As Long): End Sub
Attribute Display.VB_UserMemId = -182
Sub PrintOut(): End Sub
Attribute PrintOut.VB_UserMemId = -183
Property Get Sender() As Object: End Property
Attribute Sender.VB_UserMemId = -184
Property Get SentOnBehalfOf() As Object: End Property
Attribute SentOnBehalfOf.VB_UserMemId = -185
Property Get Session() As vbMAPI_Session: End Property
Attribute Session.VB_UserMemId = -186
Property Get Store() As vbMAPI_Store: End Property
Attribute Store.VB_UserMemId = -187
Property Get Parent() As vbMAPI_Folder: End Property
Attribute Parent.VB_UserMemId = -188
Sub Delete(Optional ByVal DeleteType As EnumDeleteType = DeleteType_MoveToDeletedItems): End Sub
Attribute Delete.VB_UserMemId = -189
Function Move(ByVal DestFolder As vbMAPI_Folder, Optional ByVal ImplicitSave As Boolean = True) As vbMAPI_MailItem: End Function
Attribute Move.VB_UserMemId = -190
Function Copy(ByVal DestFolder As vbMAPI_Folder, Optional ByVal ImplicitSave As Boolean = True) As vbMAPI_MailItem: End Function
Attribute Copy.VB_UserMemId = -191
Sub SaveAs(ByVal FilePath As String, Optional ByVal FileFormat As EnumMessageFileFormat = FileFormat_MSG): End Sub
Attribute SaveAs.VB_UserMemId = -192
Sub AbortSubmit(): End Sub
Attribute AbortSubmit.VB_UserMemId = -193
Sub MarkRead(Optional ByVal SuppressReadReceipt As Boolean = True): End Sub
Attribute MarkRead.VB_UserMemId = -194
Function Reply() As vbMAPI_MailItem: End Function
Attribute Reply.VB_UserMemId = -195
Function ReplyAll() As vbMAPI_MailItem: End Function
Attribute ReplyAll.VB_UserMemId = -196
Function Forward() As vbMAPI_MailItem: End Function
Attribute Forward.VB_UserMemId = -197
Function CastToAppointmentItem() As Object: End Function
Attribute CastToAppointmentItem.VB_UserMemId = -198
Function CastToContactItem() As Object: End Function
Attribute CastToContactItem.VB_UserMemId = -199
Sub Import(ByVal FilePath As String, ByVal ImportType As EnumImportType): End Sub
Attribute Import.VB_UserMemId = -200
