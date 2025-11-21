VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Buffer As String

Public Function InsertDWORDArray(data() As Long)
    Dim i As Integer
    For i = LBound(data) To UBound(data) Step 1
        Buffer = Buffer & MakeDWORD(data(i))
    Next i
End Function

Public Function InsertDWORD(data As Long): Buffer = Buffer & MakeDWORD(data): End Function

Public Function InsertData(data As String): Buffer = Buffer & data: End Function

Public Function InsertWORD(data As Integer): Buffer = Buffer & MakeWORD(data): End Function

Public Function InsertByte(data As Integer): Buffer = Buffer & Chr(data): End Function

Public Sub InsertBytes(data As String)
    Dim i As Long
    Dim enqueueer As String
    For i = 1 To Len(data) Step 3
        enqueueer = enqueueer & Chr(Val("&h0" & Mid(data, i, 2)))
    Next i
    Buffer = Buffer & enqueueer
End Sub

Public Function InsertNTString(data As String): Buffer = Buffer & data & Chr(0): End Function

Public Function InsertATString(data As String): Buffer = Buffer & data & Chr(&HA): End Function

Public Function InsertNonNTString(data As String): Buffer = Buffer & data: End Function

Public Function InsertNonNTStringArray(data() As String)
    Dim i As Integer
    For i = LBound(data) To UBound(data) Step 1
        Buffer = Buffer & data(i)
    Next
End Function

Public Function MakeDWORD(Value As Long) As String
    Dim Result As String * 4
    CopyMemory ByVal Result, Value, 4
    MakeDWORD = Result
End Function

Function MakeWORD(Value As Integer) As String
    Dim Result As String * 2
    CopyMemory ByVal Result, Value, 2
    MakeWORD = Result
End Function

Public Function Clear(): Buffer = vbNullString: End Function

Public Function SendPacket(p As Byte)
Dim pData As String
    pData = Chr(&HFF) & Chr(p) & MakeWORD(Len(Buffer) + 4) & Buffer
    If BNET.varAdvDebug Then
        If Not ((p = &H25 Or p = &H0) And blEnterChat = True) Then
            AddChat Color.BotInfo, svrBNCS, Color.Carrot, "[C 0x" & Hex(p) & "] ", Color.BotInfo, PNBNCS(p)
            AddChatDebug Color.Bot, DebugOutput(pData)
        End If
    End If

    If frmMain.wsBNET.State <> sckConnected Then
        Clear
        Exit Function
    End If
    frmMain.wsBNET.SendData pData
    Clear
End Function

Public Function SendRPacket(p As Byte)
Dim pData As String
    pData = MakeWORD(Len(Buffer) + 3) & Chr(p) & Buffer
    If BNET.varAdvDebug Then
        AddChat Color.BotInfo, svrMCP, Color.Carrot, "[C 0x" & Hex(p) & "] ", Color.BotInfo, PNMCP(p)
        AddChatDebug Color.Bot, DebugOutput(pData)
    End If
    If frmRealm.wsMCP.State <> sckConnected Then
        Clear
        Exit Function
    End If
    frmRealm.wsMCP.SendData pData
    Clear
End Function

Public Function SendBNLSPacket(p As Byte)
Dim pData As String
    pData = MakeWORD(Len(Buffer) + 3) & Chr(p) & Buffer
    If BNET.varAdvDebug Then
        AddChat Color.BotInfo, svrBNLS, Color.Carrot, "[C 0x" & Hex(p) & "] ", Color.BotInfo, PNBNLS(p)
        AddChatDebug Color.Bot, DebugOutput(pData)
    End If
    If frmMain.wsBNLS.State <> sckConnected Then
        Clear
        Exit Function
    End If
    frmMain.wsBNLS.SendData pData
    Clear
End Function

Public Function SendD2GSPacket(p As Byte)
Dim pData As String
    pData = Chr(p) & Buffer
    If BNET.varAdvDebug Then
        AddChat Color.BotInfo, svrD2GS, Color.Carrot, "[C 0x" & Hex(p) & "] ", Color.BotInfo, PNCD2GS(p)
        AddChatDebug Color.Bot, DebugOutput(pData)
    End If
    If frmMain.wsD2GS.State <> sckConnected Then
        Clear
        Exit Function
    End If
    frmMain.wsD2GS.SendData pData
    Clear
End Function

Public Function GetDWORD(data As String) As Long
    Dim tmp As String
    Dim A As String, B As String, c As String, d As String
    tmp = ToHex(data)
    A = Mid(tmp, 1, 2)
    B = Mid(tmp, 3, 2)
    c = Mid(tmp, 5, 2)
    d = Mid(tmp, 7, 2)
    tmp = d & c & B & A
    GetDWORD = Val("&H" & tmp)
End Function

Public Function SendServer(SOCKET As Winsock)
    SOCKET.SendData MakeWORD(Len(Buffer) + 2)
    SOCKET.SendData Buffer
    Clear
End Function
