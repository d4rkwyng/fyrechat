Attribute VB_Name = "modBCP"
Option Explicit
'DLL Calls
Declare Function LoadPlugin Lib "BCEL.dll" (ByVal lpszPlugin As String, ByRef lpPluginInfo As PluginInformation4) As Long
Declare Function FreePlugin Lib "BCEL.dll" (ByVal lpszPlugin As String) As Long
Declare Function GetMasterInterface Lib "BCEL.dll" (ByRef info As PluginInformation4) As Long
Declare Function SetMasterInterface Lib "BCEL.dll" (ByRef info As PluginInformation4) As Long
Declare Function CallPluginTimers Lib "BCEL.dll" (ByVal dwWaitTime As Long) As Long
Declare Function CallPluginMessageNotifications Lib "BCEL.dll" (ByVal dwConnectionId As Long, ByVal MessageId As Long, ByVal data As Any, ByVal size As Long) As Long
Declare Function CallPluginCommandNotifications Lib "BCEL.dll" (ByVal lpszMessage As String, ByVal lpszSendingUser As String, ByVal bFromTelnet As Long) As Long
Declare Sub CallPluginEventNotifications Lib "BCEL.dll" ()
Declare Sub SetPlatformId Lib "BCEL.dll" (ByVal Product As Long, ByVal Platform As Long)
Declare Sub SetVersionInformation Lib "BCEL.dll" (ByVal Version As Long, ByVal BuildNumber As Long)
Declare Sub CallPluginConnectNotifications Lib "BCEL.dll" (ByVal dwConnectionId As Long, ByVal dwEventCode As Long)

'Added Functions
Declare Function CStrLen Lib "kernel32" Alias "lstrlenA" (ByVal Ptr As Long) As Long
Declare Sub CStrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal dest As String, ByVal SrcPtr As Long)
Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

'Variables
Private dwColorId As Long

'BinaryChat Plugin API
Public Const NUM_INTERFACES = 3

Private Enum ConnectionType
    Connection_Local = 0
    Connection_MCP
    Connection_GS
End Enum

Private Enum ConnectionState
    Connection_Closed = 0
    Connection_Opened
End Enum

Private Enum BotWindow
    Main_Window = 0
    lstChannel
    txtChanName
    User_Input
    Chat_Output
    Account_Manager
    Realm_Manager
    Download_Manager
    News_Window
    Filters_Manager
    Channels_Manager
End Enum

Private Enum BotMenu
    Main_Window_Menu = 0
    Channel_User_Popup_Menu
    Tray_Menu
End Enum

Private Enum Output
    Output_Normal
    Output_Timestamp
End Enum

Private Enum Activation
    Plugin_Activate
    Plugin_Deactivate
End Enum

Private Enum PBCE
    PBCE_From_Telnet = &H1
End Enum

Private Enum QCE
    QCE_Fix_D2_Command = &H1
    QCE_Process_Modifiers = &H2
    QCE_Reset_Away_Idle = &H4
End Enum

Private Enum Controls
    Control_Start = 0
    Control_Stop
    Control_Hide
    Control_Show
    Control_Create_Profile
    Control_Exit
    Control_Load_Library
    Control_Free_Library
    Control_Get_Service_Status_Handle
    Control_Get_Service_Status
End Enum

Private Enum Tray
    Tray_Notify_Window = 0
End Enum

Private Enum BotOutput
    White = 0
    Gray = 1
    Blue = 2
    Green = 3
    Red = 4
    Yellow = 5
    Purple = 6
    BlueGreen = 7
    MidBlue = 8
    DarkBlue = 9
    DarkGreen = 10
    DarkYellow = 11
    DarkRed = 12
    DarkPurple = 13
    LightBlue = 14
End Enum

Private Type PluginInformation
    dwSize As Long
    PNetClearMessage As Long
    PNetIsMessageSet As Long
    PNetSetMessage As Long
    PNetQueueMessage As Long
    PNetQueueChatMessage As Long
    PSysQueryStateInformation As Long
    PSysQueryChannelInformation As Long
    PNetCloseConnection As Long
    PNetSendConnectionMessage As Long
    PUISetColor As Long
    PUIWriteTimestamp As Long
    PUIWriteString As Long
    
    hBotNetClient As Long
    'BotNetConnectionState * lpBotNetConnection
    lpBotNetConnection As Long
    
    MessageHook As Long
    MessageHookParam As Long
    
    ConnectionHook As Long
    ConnectionHookParam As Long
    
    PluginTimer As Long
    PluginTimerParam As Long
    
    CommandHook As Long
    CommandHookParam As Long
    
    hPlugin As Long
    
    EventHook As Long
    EventHookParam As Long
    hPluginEvent As Long
End Type

Private Type PluginInformation2
    Interface1 As PluginInformation
    PSysProcessBotCommand As Long
End Type

Private Type PluginInformation3
    Interface2 As PluginInformation2
    PRegGetRegistryPath As Long
    PNetResetAwayIdle As Long
    PUIGetWindow As Long
    PSysIsService As Long
    PUIGetMenu As Long
End Type
 
Private Type PluginInformation4
    Interface3 As PluginInformation3
    PFltIsUserFiltered As Long
    PFltIsMessageFiltered As Long
    PFltEnumUserFilters As Long
    PFltEnumMessageFilters As Long
    PFltAddUserFilter As Long
    PFltAddMessageFilter As Long
    PFltRemoveUserFilter As Long
    PFltRemoveMessageFilter As Long
    PUINotifyMsg As Long
    PSysEnumLoadedPlugins As Long
End Type

Private Type PluginInformation5
    Interface4 As PluginInformation4
    EncryptionInstance As Long
    RSATarget As Long
    DESTarget As Long

    PSysProcessBotCommandEx As Long
    PNetQueueChatMessageEx As Long
    PSysQueueDPC As Long

    PluginDestroy As Long
    PluginDestroyParam As Long

    OutputHook As Long
    OutputHookParam As Long
End Type

'Preliminary
Private Type PluginInformation6
    Interface5 As PluginInformation5
    MessageSendHook As Long
    MessageSendHookParam As Long

    TextModifierSendHook As Long
    TextModifierSendHookParam As Long

    TextModifierReceiveHook As Long
    TextModifierReceiveHookParam As Long

    'Add BotNetXHook to replace old-style method of interacting
    'with BotNetClient so we can unload
End Type

Private Type StateInformation
    dwSize As Long
    szChannelName(0 To 256) As Byte
    dwChannelFlags As Long
    szUniqueName(0 To 256) As Byte
    szAccountName(0 To 256) As Byte
    szStatString(0 To 256) As Byte
    dwConnectionState(0 To NUM_INTERFACES) As Long
End Type

Private Type PluginInformationEx
    dwSize As Long
    dwPluginInformationSize As Long 'Highest PluginInformation supported
    PSysGetCurrentBinaryChat As Long
    SPSysSetCurrentBinaryChat As Long
    PSysEnumBinaryChats As Long
    PSysEnumBinaryChatProfiles As Long
    PSysEnumExPlugins As Long
    PSysControlBinaryChat As Long
    PSysGetBinaryChatProfileName As Long
    PRegGetProfileRegistryPath As Long
    PUIGetExWindow As Long
    PSysPluginInformationToBinaryChat As Long
    PSysBinaryChatToPluginInformation As Long
    PSysGetBotId As Long
    PSysGetBinaryChat As Long

    PluginControl As Long
    PluginControlParam As Long

    PluginDestroyEx As Long
    PluginDestroyExParam As Long

    hPlugin As Long
    hInterfaceList As Long 'Reserved for internal use
    'End BinaryChat Plugin API
    
    'BotNetClient API (Optional)
    'End BotNetClient API
End Type

Public Function RemovePlugin(ByVal Plugin As String) As Boolean
    RemovePlugin = FreePlugin(Plugin)
End Function

Public Function AddPlugin(ByVal Plugin As String) As Boolean
On Local Error Resume Next
    Dim Interface As PluginInformation4
    Call GetMasterInterface(Interface)
    With Interface.Interface3.Interface2.Interface1
        .dwSize = Len(Interface)
        .PNetClearMessage = Address(AddressOf ClearMessageProc)
        .PNetIsMessageSet = Address(AddressOf IsMessageSetProc)
        .PNetSetMessage = Address(AddressOf SetMessageProc)
        .PNetQueueMessage = Address(AddressOf QueueMessageProc)
        .PNetQueueChatMessage = Address(AddressOf QueueChatMessageProc)
        .PSysQueryStateInformation = Address(AddressOf QueryStateInformationProc)
        .PSysQueryChannelInformation = Address(AddressOf QueryChannelInformationProc)
        .PNetCloseConnection = Address(AddressOf CloseConnectionProc)
        .PNetSendConnectionMessage = Address(AddressOf SendConnectionMessageProc)
        .PUISetColor = Address(AddressOf SetOutputColorProc)
        .PUIWriteTimestamp = Address(AddressOf WriteOutputTimestampProc)
        .PUIWriteString = Address(AddressOf WriteOutputStringProc)
        .hBotNetClient = 0
        .lpBotNetConnection = 0
    End With
    With Interface.Interface3.Interface2
        .PSysProcessBotCommand = Address(AddressOf ProcessBotCommandProc)
    End With
    With Interface.Interface3
        .PRegGetRegistryPath = Address(AddressOf GetRegistryPathProc)
        .PNetResetAwayIdle = Address(AddressOf ResetAwayIdleProc)
        .PUIGetWindow = Address(AddressOf GetWindowProc)
        .PSysIsService = Address(AddressOf IsServiceProc)
        .PUIGetMenu = Address(AddressOf GetMenuProc)
    End With
    With Interface
        .PFltIsUserFiltered = Address(AddressOf FltIsUserFiltered)
        .PFltIsMessageFiltered = Address(AddressOf FltIsMessageFiltered)
        .PFltEnumUserFilters = Address(AddressOf FltEnumUserFilters)
        .PFltEnumMessageFilters = Address(AddressOf FltEnumMessageFilters)
        .PFltAddUserFilter = Address(AddressOf FltAddUserFilter)
        .PFltAddMessageFilter = Address(AddressOf FltAddMessageFilter)
        .PFltRemoveUserFilter = Address(AddressOf FltRemoveUserFilter)
        .PFltRemoveMessageFilter = Address(AddressOf FltRemoveMessageFilter)
        .PUINotifyMsg = Address(AddressOf UINotifyMsg)
        .PSysEnumLoadedPlugins = Address(AddressOf SysEnumLoadedPlugins)
    End With

    'With Interface
    '    .PSysProcessBotCommandEx = Address(AddressOf ProcessBotCommandExProc)
    '    .PNetQueueChatMessageEx = Address(AddressOf QueueChatMessageExProc)
    '    .PSysQueueDPC = Address(AddressOf QueueDPCProc)
    '    .PluginDestroy = Address(AddressOf PluginDestroyProc)
    '    .OutputHook = Address(AddressOf OutputHookProc)
    'End With
    'With Interface
    '    .MessageSendHook = Address(AddressOf MessageSendHookProc)
    '    .TextModifierSendHook = Address(AddressOf TextModifierSendHookProc)
    '    .TextModifierReceiveHook = Address(AddressOf TextModifierReceiveHookProc)
    'End With

    Call SetMasterInterface(Interface)
    AddPlugin = LoadPlugin(Plugin, Interface)
End Function

Public Function AddPluginEx(ByVal Plugin As String) As Boolean
On Local Error Resume Next
    Dim Interface As PluginInformationEx
    'Call GetMasterInterface(Interface) 'GetMasterInterfaceEx?
    With Interface
        .dwSize = Len(Interface)
        .dwPluginInformationSize = 4
        .PSysGetCurrentBinaryChat = Address(AddressOf GetCurrentBinaryChatProc)
        .SPSysSetCurrentBinaryChat = Address(AddressOf SetCurrentBinaryChatProc)
        .PSysEnumBinaryChats = Address(AddressOf EnumBinaryChatsProc)
        .PSysEnumBinaryChatProfiles = Address(AddressOf EnumBinaryChatProfilesProc)
        .PSysEnumExPlugins = Address(AddressOf EnumLoadedPluginsProc)
        .PSysControlBinaryChat = Address(AddressOf ControlBinaryChatProc)
        .PSysGetBinaryChatProfileName = Address(AddressOf GetBinaryChatProfileNameProc)
        .PRegGetProfileRegistryPath = Address(AddressOf GetProfileRegistryPathProc)
        .PUIGetExWindow = Address(AddressOf ExGetWindowProc)
        .PSysPluginInformationToBinaryChat = Address(AddressOf PluginInformationToBinaryChatProc)
        .PSysBinaryChatToPluginInformation = Address(AddressOf BinaryChatToPluginInformationProc)
        .PSysGetBotId = Address(AddressOf GetBotIdProc)
        .PSysGetBinaryChat = Address(AddressOf GetBinaryChatProc)
        .PluginControl = Address(AddressOf PluginControlProc)
        .PluginDestroyEx = Address(AddressOf PluginDestroyExProc)
    End With
    'Call SetMasterInterface(Interface)
    'AddPlugin = LoadPlugin(Plugin, Interface)
End Function

'Interface 1
Private Function ClearMessageProc(ByVal MessageId As Long) As Long
    ClearMessageProc = 1&
End Function

Private Function IsMessageSetProc(ByVal MessageId As Long) As Long
    IsMessageSetProc = 1&
End Function

Private Function SetMessageProc(ByVal MessageId As Long) As Long
    SetMessageProc = 1&
End Function

Private Function QueueMessageProc(ByVal MessageId As Long, _
    ByVal lpMessageData As Long, ByVal nMessageSize As Long) As Long
    Dim MessageData As String, MsgId As Byte
    MsgId = MessageId
    MessageData = String(nMessageSize, vbNullChar)
    Call CopyMemory(ByVal MessageData, ByVal lpMessageData, nMessageSize)
    sPB.InsertNonNTString MessageData
    sPB.SendPacket MsgId
    QueueMessageProc = 1&
End Function

Private Function QueueChatMessageProc(ByVal lpszChatMessage As Long, _
    ByVal bFixD2Command As Long) As Long
    Dim MessageData As String, SplitData() As String, tmpData As String
    Dim i As Integer
    Dim Queue As ListBox
    Set Queue = frmMain.lstEnqueue
    MessageData = Space$(CStrLen(lpszChatMessage))
    Call CStrCpy(MessageData, lpszChatMessage)
    If bFixD2Command = 1 Then
        SplitData() = Split(MessageData, " ")
        If Not Mid(SplitData(1), 1, 1) = "*" Then
            tmpData = ""
            SplitData(1) = "*" & SplitData(1)
            For i = 0 To Len(SplitData(i))
                tmpData = tmpData & SplitData(i)
            Next i
            MessageData = tmpData
        End If
    End If
    frmMain.User_Input = MessageData
    frmMain.User_Input_KeyPress (13)
    QueueChatMessageProc = 1&
End Function

Private Function QueryStateInformationProc(ByVal lpStateInfo As Long) As Long
    Dim State As StateInformation
    QueryStateInformationProc = 1&
End Function

Private Function QueryChannelInformationProc(ByVal lpCallback As Long, _
    ByVal lParam As Long) As Long
    QueryChannelInformationProc = 1&
End Function

Private Function CloseConnectionProc(ByVal dwConnectionId As Long) As Long
    Select Case dwConnectionId
        Case 0: frmMain.wsBNET.Close
        Case 1: frmRealm.wsMCP.Close
        Case 2: AddChat Color.BotError, "No GameServer Support"
        Case Else: AddChat Color.BotError, "Error in CloseConnectionProc: Unknown Connection ID"
    End Select
    CloseConnectionProc = 1&
End Function

Private Function SendConnectionMessageProc(ByVal dwConnectionId As Long, _
    ByVal MessageId As Long, ByVal lpMessageData As Long, _
    ByVal nMessageSize As Long) As Long
    Dim MessageData As String, MsgId As Byte
    MsgId = MessageId
    MessageData = String(nMessageSize, vbNullChar)
    Call CopyMemory(ByVal MessageData, ByVal lpMessageData, nMessageSize)
    Select Case dwConnectionId
        Case 0: Send MessageData, frmMain.wsBNET
        Case 1: sPB.SendRPacket MsgId
        Case 2: AddChat Color.BotError, "No GameServer Support"
        Case Else: AddChat Color.BotError, "Error in SendConnectionMessageProc: Unknown Connection ID"
    End Select
    SendConnectionMessageProc = 1&
End Function

Private Function SetOutputColorProc(ByVal ColorId As BotOutput) As Long
     Select Case ColorId
         Case White: dwColorId = RGB(&HFF, &HFF, &HFF)
         Case Gray: dwColorId = RGB(&H7F, &H7F, &H7F)
         Case Blue: dwColorId = RGB(&H0, &HFF, &HFF)
         Case Green: dwColorId = RGB(&H32, &HFF, &H32)
         Case Red: dwColorId = RGB(&HFF, &H32, &H32)
         Case Yellow: dwColorId = RGB(&HFF, &HFF, &H0)
         Case Purple: dwColorId = RGB(&HFF, &H0, &HFF)
         Case BlueGreen: dwColorId = RGB(&H0, &H7F, &H7F)
         Case MidBlue: dwColorId = RGB(&H0, &H0, &HFF)
         Case DarkBlue: dwColorId = RGB(&H0, &H0, &H7F)
         Case DarkGreen: dwColorId = RGB(&H10, &HAF, &H10)
         Case DarkYellow: dwColorId = RGB(&H7F, &H7F, &H0)
         Case DarkRed: dwColorId = RGB(&H7F, &H0, &H0)
         Case DarkPurple: dwColorId = RGB(&H7F, &H0, &H7F)
         Case LightBlue: dwColorId = RGB(&H0, &H6F, &HFF)
         Case Else: dwColorId = RGB(&HFF, &HFF, &HFF)
     End Select
     SetOutputColorProc = 1&
 End Function
 
Private Function WriteOutputTimestampProc() As Long
    Dim timestamp As String
    timestamp = "[" & Format(Time, "hh:mm:ss") & "] "
    With frmMain.Chat_Output
        .SelStart = Len(.text)
        .SelLength = 0
        .SelColor = Color.timestamp
        .SelText = timestamp
    End With
    WriteOutputTimestampProc = 1&
 End Function
 
Private Function WriteOutputStringProc(ByVal lpszString As Long) As Long
    Dim MessageData As String
    MessageData = String(CStrLen(lpszString), vbNullChar)
    Call CStrCpy(MessageData, lpszString)
    With frmMain.Chat_Output
        .SelLength = 0
        .SelColor = dwColorId
        .SelText = MessageData
        .SelStart = Len(.text)
    End With
    WriteOutputStringProc = 1&
End Function

'Interface 2
Private Function ProcessBotCommandProc(ByVal lpszCommand As Long) As Long
    Dim MessageData As String
    MessageData = String(CStrLen(lpszCommand), vbNullChar)
    Call CStrCpy(MessageData, lpszCommand)
    frmMain.User_Input = MessageData
    frmMain.User_Input_KeyPress (13)
    ProcessBotCommandProc = 1&
End Function

'Interface 3
Private Function GetRegistryPathProc() As Long
    Dim strRegLoc As String, tmpLoc As String
    Dim Reg As clsRegistry
    tmpLoc = App.Path
    If InStr(tmpLoc, "\") Then tmpLoc = Replace(tmpLoc, "\", "/")
    strRegLoc = "Software\OpenFyre\FyreChat\" & tmpLoc
    Set Reg = New clsRegistry
    Reg.hkey = HKEY_Local_MACHINE
    Reg.KeyRoot = strRegLoc
    Reg.Subkey = "Plugins"
    If Not Reg.KeyExists Then Reg.CreateKey
    
    tmpLoc = "Software\OpenFyre\FyreChat\" & tmpLoc & "\Plugins\"
    tmpLoc = StrConv(tmpLoc, vbFromUnicode)
    GetRegistryPathProc = StrPtr(tmpLoc)
End Function

Private Function ResetAwayIdleProc() As Long
    Send "/away", frmMain.wsBNET
    awayFlag = False
    tFlag = False
    LastTalk = GetTickCount()
    ResetAwayIdleProc = 1&
End Function

Private Function GetWindowProc(ByVal dwWindowId As BotWindow) As Long
    Select Case dwWindowId
        Case Main_Window: GetWindowProc = frmMain.hWnd
        Case lstChannel: GetWindowProc = frmMain.lstChannel.hWnd
        Case txtChanName: GetWindowProc = frmMain.txtChanName.hWnd
        Case User_Input: GetWindowProc = frmMain.User_Input.hWnd
        Case Chat_Output: GetWindowProc = frmMain.Chat_Output.hWnd
        Case Account_Manager: GetWindowProc = 0&
        Case Realm_Manager: GetWindowProc = frmRealm.hWnd
        Case Download_Manager: GetWindowProc = 0&
        Case Filters_Manager: GetWindowProc = 0&
        Case News_Window: GetWindowProc = 0&
        Case Channels_Manager: GetWindowProc = 0&
        Case Else: GetWindowProc = 0&
    End Select
End Function

Private Function IsServiceProc() As Long
    'True: 1, False: 0
    IsServiceProc = 1&
End Function

Private Function GetMenuProc(ByVal dwMenuId As BotMenu) As Long
    Select Case dwMenuId
        Case Main_Window_Menu: GetMenuProc = GetMenu(frmMain.hWnd)
        Case Channel_User_Popup_Menu: GetMenuProc = GetMenu(frmMain.hWnd)
        Case Tray_Menu: GetMenuProc = GetMenu(frmMain.hWnd)
        Case Else: GetMenuProc = 0&
    End Select
End Function

'Interface 4
Private Function FltIsUserFiltered(ByVal lpszUser As Long) As Long
    FltIsUserFiltered = 1&
End Function

Private Function FltIsMessageFiltered(ByVal lpszMessage As Long) As Long
    FltIsMessageFiltered = 1&
End Function

Private Function FltEnumUserFilters(ByVal lpCallback As Long, _
    ByVal lParam As Long) As Long
    FltEnumUserFilters = 1&
End Function

Private Function FltEnumMessageFilters(ByVal lpCallback As Long, _
    ByVal lParam As Long) As Long
    FltEnumMessageFilters = 1&
End Function

Private Function FltAddUserFilter(ByVal lpszFilter As Long, _
    ByVal bSaveFilters As Long) As Long
    'Dim MessageData As String
    'MessageData = String(CStrLen(lpszFilter), vbNullChar)
    'Call CStrCpy(MessageData, lpszFilter)
    'Call frmFilters.AddFilter(MessageData, frmFilters.lstFilters)
    'If bSaveFilters = 1 Then frmFilters.UpdateFilters
    FltAddUserFilter = 1&
End Function

Private Function FltAddMessageFilter(ByVal lpszMessageFilter As Long, _
    ByVal bSaveFilters As Long) As Long
    'Dim MessageData As String
    'MessageData = String(CStrLen(lpszMessageFilter), vbNullChar)
    'Call CStrCpy(MessageData, lpszMessageFilter)
    'Call frmFilters.AddFilter(MessageData, frmFilters.lstFiltMsgs)
    'If bSaveFilters = 1 Then frmFilters.UpdateFilters
    FltAddMessageFilter = 1&
End Function

Private Function FltRemoveUserFilter(ByVal lpszFilter As Long, _
    ByVal bSaveFilters As Long) As Long
    'Dim MessageData As String
    'MessageData = String(CStrLen(lpszFilter), vbNullChar)
    'Call CStrCpy(MessageData, lpszFilter)
    'Call frmFilters.DeleteFilter(MessageData, frmFilters.lstFilters)
    'If bSaveFilters = 1 Then frmFilters.UpdateFilters
    FltRemoveUserFilter = 1&
End Function

Private Function FltRemoveMessageFilter(ByVal lpszFilter As Long, _
    ByVal bSaveFilters As Long) As Long
    'Dim MessageData As String
    'MessageData = String(CStrLen(lpszFilter), vbNullChar)
    'Call CStrCpy(MessageData, lpszFilter)
    'Call frmFilters.DeleteFilter(MessageData, frmFilters.lstFiltMsgs)
    'If bSaveFilters = 1 Then frmFilters.UpdateFilters
    FltRemoveMessageFilter = 1&
End Function

Private Function UINotifyMsg(ByVal bColor As Long, _
    ByVal lpszUser As Long, ByVal lpszMessageFormat As Long) As Long
    'Cannot Work in VB
    Dim UserData As String, MessageFormat As String
    UserData = String(CStrLen(lpszUser), vbNullChar)
    MessageFormat = String(CStrLen(lpszMessageFormat), vbNullChar)
    AddChat Color.BotInfo, "[UINotifyMsg] ", Color.Bot, UserData & " triggered " & MessageFormat & "."
    UINotifyMsg = 1&
End Function

Private Function SysEnumLoadedPlugins(ByVal lpCallback As Long, _
    ByVal lParam As Long) As Long
    SysEnumLoadedPlugins = 1&
End Function

'Interface 5
Private Function ProcessBotCommandExProc(ByVal lpszMessage As Long, _
    ByVal dwFlags As Long, ByVal lpszRemoteUser As Long) As Long
    ProcessBotCommandExProc = 1&
End Function

Private Function QueueChatMessageExProc(ByVal lpszMessage As Long, _
    ByVal dwFlags As Long) As Long
    QueueChatMessageExProc = 1&
End Function

Private Function QueueDPCProc(ByVal DPC As Long, ByVal lParam As Long) As Long
    'DBC As DeferredProcedureCallProc
    QueueDPCProc = 1&
End Function

Private Function PluginDestroyProc(ByVal lParam As Long) As Long
    PluginDestroyProc = 1&
End Function

Private Function OutputHookProc(ByVal dwReason As Long, ByVal Color As Long, _
    ByVal lpszMessage As Long, ByVal lParam As Long) As Long
    OutputHookProc = 1&
End Function

'Interface 6 (Preliminary)
Private Function MessageSendHookProc(ByVal dwConnectionId As Long, _
    ByVal MessageId As Long, ByVal lpMessageData As Long, ByVal nMessageSize As Long, _
    ByVal lParam As Long) As Long
    MessageSendHookProc = 1&
End Function

Private Function TextModifierSendHookProc(ByVal dwConnectionId As Long, _
    ByVal lpszMessageText As Long, ByVal lParam As Long) As Long
    TextModifierSendHookProc = 1&
End Function

Private Function TextModifierReceiveHookProc(ByVal dwConnectionId As Long, _
    ByVal lpszMessageText As Long, ByVal lParam As Long) As Long
    TextModifierReceiveHookProc = 1&
End Function

'BCPX
Private Function GetCurrentBinaryChatProc() As Long
    GetCurrentBinaryChatProc = 1&
End Function

Private Function SetCurrentBinaryChatProc(ByVal hBinaryChat As Long) As Long
    SetCurrentBinaryChatProc = 1&
End Function

Private Function EnumBinaryChatsProc(ByVal lpCallback As Long, _
    ByVal lParam As Long) As Long
    EnumBinaryChatsProc = 1&
End Function

Private Function EnumBinaryChatProfilesProc(ByVal lpCallback As Long, _
    ByVal lParam As Long) As Long
    EnumBinaryChatProfilesProc = 1&
End Function

Private Function EnumLoadedPluginsProc(ByVal lpCallback As Long, _
    ByVal lParam As Long) As Long
    EnumLoadedPluginsProc = 1&
End Function

Private Function ControlBinaryChatProc(ByVal hBinaryChat As Long, _
    ByVal dwControlCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    ControlBinaryChatProc = 1&
End Function

Private Function GetBinaryChatProfileNameProc(ByVal hBinaryChat As Long) As Long
    GetBinaryChatProfileNameProc = 1&
End Function

Private Function GetProfileRegistryPathProc() As Long
    GetProfileRegistryPathProc = 1&
End Function

Private Function ExGetWindowProc(ByVal dwWindowId As Long) As Long
    ExGetWindowProc = 1&
End Function

Private Function PluginInformationToBinaryChatProc(ByVal lpPluginInformation As Long, _
    ByVal hInterfaceList As Long) As Long
    'PluginInformationCurrent * lpPluginInformation
    PluginInformationToBinaryChatProc = 1&
End Function

Private Function BinaryChatToPluginInformationProc(ByVal hBinaryChat As Long, _
    ByVal hInterfaceList As Long) As Long
    BinaryChatToPluginInformationProc = 1&
End Function

Private Function GetBotIdProc(ByVal hBinaryChat As Long) As Long
    GetBotIdProc = 1&
End Function

Private Function GetBinaryChatProc(ByVal dwBotId As Long) As Long
    GetBinaryChatProc = 1&
End Function

Private Function PluginControlProc(ByVal dwOperation As Long, _
    ByVal lpInterfaceInfo As Long, ByVal hBinaryChat As Long, _
    ByVal lParam As Long) As Long
    'PluginInformation * lpInterfaceInfo
    PluginControlProc = 1&
End Function

Private Function PluginDestroyExProc(ByVal lParam As Long) As Long
    PluginDestroyExProc = 1&
End Function

'Functions
Public Function Address(ByVal dwAddress As Long) As Long
     Address = dwAddress
End Function
