VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private Buffer     As String
Private Position   As Long
Private InNOut     As Boolean
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                                                     ByRef Source As Any, _
                                                                     ByVal numBytes As Long)
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Sub Class_Initialize()
    On Error Resume Next
    Position = 1
    On Error GoTo 0
End Sub

Public Sub Clear()
    On Error Resume Next
    Buffer = vbNullString
    Position = 1
    On Error GoTo 0
End Sub

Public Function GetBoolean() As Boolean
    Dim lngTemp As Long
    On Error Resume Next
    If LenBuffer() - (Position + 1) >= 4 Then
        CopyMemory lngTemp, ByVal Mid$(Buffer, Position, 4), 4
        Position = Position + 4
    Else
        Exit Function
    End If
    If lngTemp <> 0 Then
        GetBoolean = True
    End If
    On Error GoTo 0
End Function

Public Function GetBuffer() As String
    On Error Resume Next
    GetBuffer = Buffer
    On Error GoTo 0
End Function

Public Function GetByte() As Byte
    Dim bytTemp As Byte
    On Error Resume Next
        CopyMemory bytTemp, ByVal Mid$(Buffer, Position, 1), 1
        Position = Position + 1
        GetByte = bytTemp
    On Error GoTo 0
End Function

Public Function GetDWORD() As Long
    Dim lngTemp As Long
    On Error Resume Next
    CopyMemory lngTemp, ByVal Mid$(Buffer, Position, 4), 4
    Position = Position + 4
    GetDWORD = lngTemp
    On Error GoTo 0
End Function

Public Function GetDWORDArray(ByVal Amount As Long) As Long()
    Dim RtnAry() As Long
    Dim i        As Long
    On Error Resume Next
    If Not Amount < 1 Then
        If LenBuffer() - (Position + 1) >= Amount * 4 Then
            ReDim RtnAry(Amount) As Long
            For i = 0 To Amount - 1
                RtnAry(i) = GetDWORD()
            Next i
            GetDWORDArray = RtnAry
        End If
        On Error GoTo 0
    End If
End Function

Public Function GetFixedString(ByVal Length As Long) As String
    On Error Resume Next
    If LenBuffer() - (Position + 1) >= Length Then
        GetFixedString = Mid$(Buffer, Position, Length)
        Position = Position + Length
    Else
        Position = Position + Length
    End If
    On Error GoTo 0

End Function

Public Function GetMode() As Boolean
    On Error Resume Next
    GetMode = InNOut
    On Error GoTo 0
End Function

Public Function GetNonNTString(ByVal Length As Long) As String
    On Error Resume Next
    If LenBuffer() - (Position + 1) >= Length Then
        GetNonNTString = Mid$(Buffer, Position, Length)
        Position = Position + Length
    End If
    On Error GoTo 0
End Function

Public Function GetPosition() As Long

    On Error Resume Next
    GetPosition = Position
    On Error GoTo 0

End Function

Public Function GetRaw(ByVal Length As Long) As String
    On Error Resume Next
    GetRaw = Mid$(Buffer, Position, Length)
    Position = Position + Length
    On Error GoTo 0
End Function

Public Function GetString() As String
    Dim pos As Long
    On Error Resume Next
    GetString = vbNullString
    pos = InStr(Position, Buffer, Chr$(&H0))
    If pos = 0 Then
        Exit Function
    End If
    GetString = Mid$(Buffer, Position, pos - Position)
    Position = pos + 1
    On Error GoTo 0
End Function

Public Function GetStringArray(ByVal Amount As Long) As String()
    Dim RtnAry() As String
    Dim i        As Long
    On Error Resume Next
    If Not Amount < 1 Then
    ReDim RtnAry(Amount) As String
        For i = 0 To Amount - 1
            RtnAry(i) = GetString()
        Next i
        GetStringArray = RtnAry
        On Error GoTo 0
    End If
End Function

Public Function GetTermString(ByVal Terminator As Integer) As String
    Dim pos As Long
    On Error Resume Next
    GetTermString = vbNullString
    pos = InStr(Position, Buffer, Chr$(Terminator))
    If pos = 0 Then
        Exit Function
    End If
    GetTermString = Mid$(Buffer, Position, pos - Position)
    Position = pos + 1
    On Error GoTo 0

End Function

Public Function GetWORD() As Long
    Dim lngTemp As Long
    On Error Resume Next
    If LenBuffer() - (Position + 1) >= 2 Then
        CopyMemory lngTemp, ByVal Mid$(Buffer, Position, 2), 2
        Position = Position + 2
        GetWORD = lngTemp
    End If
    On Error GoTo 0
End Function

Public Function GetWORDArray(ByVal Amount As Long) As Long()
    Dim RtnAry() As Long
    Dim i        As Long
    If Not Amount < 1 Then
        If LenBuffer() - (Position + 1) >= Amount * 2 Then
            ReDim RtnAry(Amount) As Long
            For i = 0 To Amount - 1
                RtnAry(i) = GetWORD()
            Next i
            GetWORDArray = RtnAry
        Else
        End If
    End If
End Function

Public Function GetFileTime() As String
    Dim FT As FILETIME
    Dim Temp As String * 8
    Temp = Mid$(Buffer, Position, 8)
    CopyMemory FT, ByVal Temp, 8
    Position = Position + 8
    GetFileTime = FT.dwHighDateTime & Space$(1) & FT.dwLowDateTime
 End Function

Public Sub InsertByte(ByVal data As Byte)
    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    Buffer = Buffer & MakeByte(data)
    On Error GoTo 0
End Sub

Public Sub InsertDWORD(ByVal data As Long)
    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    Buffer = Buffer & MakeDWORD(data)
    On Error GoTo 0
End Sub

Public Sub InsertDWORDArray(data() As Long)
    Dim i As Long
    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    If UBound(data) = 0 Then
        If data(0) = 0 Then
            Exit Sub
        End If

    End If
    For i = LBound(data) To UBound(data)
        InsertDWORD data(i)
    Next i
    On Error GoTo 0
End Sub

Public Sub InsertNonNTString(ByVal data As String)
    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    Buffer = Buffer & data
    On Error GoTo 0
End Sub

Public Sub InsertString(ByVal data As String)
    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    Buffer = Buffer & data & Chr$(&H0)
    On Error GoTo 0

End Sub

Public Sub InsertStringArray(data() As String)
    Dim i As Long
    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    If UBound(data) = 0 Then
        If Len(data(0)) = 0 Then
            Exit Sub
        End If
    End If
    For i = LBound(data) To UBound(data)
        InsertString data(i)
    Next i
    On Error GoTo 0
End Sub

Public Sub InsertStringList(data() As String)
    Dim i As Long
    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    If UBound(data) = 0 Then
        If Len(data(0)) = 0 Then
            Exit Sub
        End If
    End If
    For i = LBound(data) To UBound(data)
        InsertString data(i)
    Next i
    InsertByte &H0
    On Error GoTo 0
End Sub

Public Sub InsertWORD(ByVal data As Integer)
    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    Buffer = Buffer & MakeWORD(data)
    On Error GoTo 0
End Sub

Public Sub InsertWORDArray(data() As Integer)
    Dim i As Long
    On Error Resume Next
    If InNOut Then
        Buffer = vbNullString
        InNOut = False
    End If
    If UBound(data) = 0 Then
        If data(0) = 0 Then
            Exit Sub
        End If
    End If
    For i = LBound(data) To UBound(data)
        InsertWORD data(i)
    Next i
    On Error GoTo 0
End Sub

Public Function LenBuffer() As Long
    On Error Resume Next
    LenBuffer = Len(Buffer)
    On Error GoTo 0
End Function

Public Function MakeByte(ByVal data As Byte) As String
    Dim Result As String * 1
    On Error Resume Next
    CopyMemory ByVal Result, data, 1
    MakeByte = Result
    On Error GoTo 0
End Function

Public Function MakeDWORD(ByVal data As Long) As String
    Dim Result As String * 4
    On Error Resume Next
    CopyMemory ByVal Result, data, 4
    MakeDWORD = Result
    On Error GoTo 0
End Function

Public Function MakeWORD(ByVal data As Integer) As String
    Dim Result As String * 2
    On Error Resume Next
    CopyMemory ByVal Result, data, 2
    MakeWORD = Result
    On Error GoTo 0
End Function

Public Sub SendRaw(SOCKET As Winsock)
    On Error Resume Next
    If Not InNOut Then
        SOCKET.SendData Buffer
        Clear
        On Error GoTo 0
    End If
End Sub

Public Sub SetBuffer(ByVal data As String)
    On Error Resume Next
    Buffer = data
    InNOut = True
    On Error GoTo 0
End Sub

Public Sub SetInNOut(ByVal Bool As Boolean)
    On Error Resume Next
    InNOut = Bool
    On Error GoTo 0
End Sub

Public Sub SetPosition(ByVal pos As Long): Position = pos: End Sub

Public Sub Skip(ByVal Bytes As Long)
    On Error Resume Next
    Position = Position + Bytes
    On Error GoTo 0
End Sub




