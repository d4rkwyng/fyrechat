VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Battle.net Events
Private Const ID_USER = &H1                'User in channel
Private Const ID_JOIN = &H2                'User joined channel
Private Const ID_LEAVE = &H3               'User left channel
Private Const ID_WHISPFROM = &H4           'Recieved whisper
Private Const ID_TALK = &H5                'Chat text
Private Const ID_BROADCAST = &H6           'Server broadcast
Private Const ID_CHAN = &H7                'Channel information
Private Const ID_FLAGS = &H9               'Flags update
Private Const ID_WHISPTO = &HA             'Sent whisper
Private Const ID_CHANNELFULL = &HD         'Channel full
Private Const ID_CHANNELDOESNOTEXIST = &HE 'Channel doesn't exist
Private Const ID_CHANNELRESTRICTED = &HF   'Channel is restricted
Private Const ID_INFO = &H12               'Information
Private Const ID_ERROR = &H13              'Error message
Private Const ID_EMOTE = &H17              'Emote

'Battle.net Events Defined
Public Event OnUser(ByVal username As String, ByVal flags As Long, ByVal Message As String, ByVal ping As Long)
Public Event OnJoin(ByVal username As String, ByVal flags As Long, ByVal Message As String, ByVal ping As Long)
Public Event OnLeave(ByVal username As String, ByVal flags As Long, ByVal ping As Long)
Public Event OnWhisperFrom(ByVal username As String, ByVal flags As Long, ByVal Message As String, ByVal ping As Long)
Public Event OnTalk(ByVal username As String, ByVal flags As Long, ByVal Message As String, ByVal ping As Long)
Public Event OnChannel(ByVal ChannelName As String, ByVal flags As Long)
Public Event OnFlags(ByVal username As String, ByVal flags As Long, ByVal Message As String, ByVal ping As Long)
Public Event OnWhisperTo(ByVal username As String, ByVal flags As Long, ByVal Message As String, ByVal ping As Long)
Public Event OnInfo(ByVal Message As String)
Public Event OnError(ByVal Message As String)
Public Event OnEmote(ByVal username As String, ByVal flags As Long, ByVal Message As String, ByVal ping As Long)
Public Event OnUnknown(ByVal UnknownString As String)

Public Sub DispatchMessage(ByVal DataBuf As String)
    Dim username As String, Message As String
    Dim EID As Long, ping As Long, flags As Long

    EID = MakeLong(Mid$(DataBuf, 5, 4))
    flags = MakeLong(Mid$(DataBuf, 9, 4))
    ping = MakeLong(Mid$(DataBuf, 13, 4))
    username = KillNull(Mid$(DataBuf, 29))
    Message = KillNull(Mid$(DataBuf, Len(username) + 30))

    If (Diablo2) And InStr(username, "*") Then
        Dim spt() As String
        spt() = Split(username, "*", 2)
         If UBound(spt()) <= 0 Then
            username = username
        Else
            username = spt(1)
        End If
    End If

    Select Case EID
        Case ID_TALK
            RaiseEvent OnTalk(username, flags, Message, ping)
            Exit Sub
        Case ID_EMOTE
            RaiseEvent OnEmote(username, flags, Message, ping)
            Exit Sub
        Case ID_CHAN
            If OnChan Then AddChat Color.Self, "Entered chat: " & BNET.TrueUsername & " (" & ping & "ms), using: " & ProductName(BNET.Product) & "."
            RaiseEvent OnChannel(Message, flags)
            Exit Sub
        Case ID_USER
            RaiseEvent OnUser(username, flags, Message, ping)
            Exit Sub
        Case ID_JOIN
            RaiseEvent OnJoin(username, flags, Message, ping)
            Exit Sub
        Case ID_LEAVE
            RaiseEvent OnLeave(username, flags, ping)
            Exit Sub
        Case ID_WHISPTO
            RaiseEvent OnWhisperTo(username, flags, Message, ping)
            Exit Sub
        Case ID_WHISPFROM
            RaiseEvent OnWhisperFrom(username, flags, Message, ping)
            Exit Sub
        Case ID_INFO
            RaiseEvent OnInfo(Message)
            Exit Sub
        Case ID_FLAGS
            RaiseEvent OnFlags(username, flags, Message, ping)
            Exit Sub
        Case ID_BROADCAST
            RaiseEvent OnInfo(Message)
            Exit Sub
        Case ID_ERROR
            RaiseEvent OnError(Message)
            Exit Sub
        Case Else
            RaiseEvent OnUnknown(Message)
            Exit Sub
    End Select
End Sub
