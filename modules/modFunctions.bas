Attribute VB_Name = "modFunctions"
Option Explicit
Public Enum WildCard
    None = 0
    AsteriskOnRight = 1
    AsteriskOnLeft = 2
    AsteriskOnBoth = 3
End Enum

Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long

Private Reg As clsRegistry

Public Function LoadConfig()
On Local Error Resume Next
    Dim strRegLoc As String, tmpLoc As String
    tmpLoc = App.Path
    If InStr(tmpLoc, "\") Then tmpLoc = Replace(tmpLoc, "\", "/")
    strRegLoc = "Software\OpenFyre\FyreChat\" & tmpLoc
    Set Reg = New clsRegistry
    Reg.hkey = HKEY_Local_MACHINE
    Reg.KeyRoot = strRegLoc
    Reg.Subkey = "Config"
    If Not Reg.KeyExists Then
        Reg.CreateKey
        FirstRun = True
    Else
        FirstRun = False
    End If
    
    BNET.username = Reg.GetRegistryValue("Username", "FyreChat")
    BNET.password = Reg.GetRegistryValue("Password", "")
    BNET.BNCSServer = Reg.GetRegistryValue("Server", "useast.battle.net")
    BNET.BNLSServer = Reg.GetRegistryValue("BNLS Server", "bnls.net")
    BNET.Product = Reg.GetRegistryValue("Product", "PXES")
    BNET.CDKey = Reg.GetRegistryValue("CDKey")
    BNET.CDKey2 = Reg.GetRegistryValue("CDKey2")
    BNET.CDKeyOwner = Reg.GetRegistryValue("Owner", "FyreChat")
    BNET.HomeChannel = Reg.GetRegistryValue("Channel", "Op FyreChat")
    BNET.varCRealm = Reg.GetRegistryValue("Realm AutoLogin", 0)
    BNET.Character = Reg.GetRegistryValue("Character", "")
    BNET.Email = Reg.GetRegistryValue("Account Email", "")
    BNET.varAutoCon = Reg.GetRegistryValue("AutoConnect", 0)
    BNET.varARCon = Reg.GetRegistryValue("Reconnect", 0)
    BNET.varLagPlug = Reg.GetRegistryValue("LagPlug", "None")
    BNET.varUDP = Reg.GetRegistryValue("UDP", 0)
    BNET.varProdSpec = Reg.GetRegistryValue("Join Product", 0)
    BNET.varEnterLeave = Reg.GetRegistryValue("Join/Leave Notifications", 1)
    BNET.varStatusUpdate = Reg.GetRegistryValue("Status Update", 1)
    BNET.varPingFlag = Reg.GetRegistryValue("PingFlags", 0)
    BNET.varLog = Reg.GetRegistryValue("Logging", 0)
    BNET.varDebugMode = Reg.GetRegistryValue("Debug Mode", 0)
    BNET.varAdvDebug = Reg.GetRegistryValue("Advanced Debug Mode", 0)
    BNET.varShowSvr = Reg.GetRegistryValue("Show Server Types", 1)
    BNET.WhispWin = Reg.GetRegistryValue("Whisper Window", 1)
    BNET.varNews = Reg.GetRegistryValue("Show News", 0)
    BNET.varSysTray = Reg.GetRegistryValue("MinToTray", 0)
    BNET.varConfirmExit = Reg.GetRegistryValue("Exit Prompt", 1)
    BNET.FloodProt = Reg.GetRegistryValue("Flood Protection", 1)
    BNET.MaxLength = Reg.GetRegistryValue("Max Chat Length", "10000")
    BNET.varAwayIdle = Reg.GetRegistryValue("Away Idle", 1)
    BNET.IdleInt = Reg.GetRegistryValue("Idle Interval", "1200")
    BNET.Idle = Reg.GetRegistryValue("Idle Text", "/me - FyreChat %ver - Uptime: %uptime")
    BNET.varIdle = Reg.GetRegistryValue("Normal Idle", 0)
    BNET.varCountIdle = Reg.GetRegistryValue("Hmm Idle", 0)
    BNET.Fonts = Reg.GetRegistryValue("Fonts", "Arial")
    BNET.FSize = Reg.GetRegistryValue("Font Size", "9")
    BNET.BGColors = Reg.GetRegistryValue("BackGround Color", "Black")
    
    'Weblist
    blWL_Enabled = Reg.GetRegistryValue("WebList Enabled", 0)
    strWL_FTPAddress = Reg.GetRegistryValue("FTP Address", "domain.com")
    strWL_FTPLoc = Reg.GetRegistryValue("FTP Location", "public_html/")
    strWL_FTPUser = Reg.GetRegistryValue("FTP User", "Anonymous")
    strWL_FTPPass = Reg.GetRegistryValue("FTP Password", "")
    intWL_UpTimer = Reg.GetRegistryValue("FTP Upload Interval", 60)
    
    Reg.Subkey = "Colors"
    If Not Reg.KeyExists Then Reg.CreateKey
    Color.Self = Reg.GetRegistryValue("Self", &HFFFF00)
    Color.Op = Reg.GetRegistryValue("Op", &H80000005)
    Color.user = Reg.GetRegistryValue("User", &HFFFF&)
    Color.Message = Reg.GetRegistryValue("Message", &H80000005)
    Color.Carrot = Reg.GetRegistryValue("Carrot", &HFFFF&)
    Color.PingFlags = Reg.GetRegistryValue("PingFlags", &H808080)
    Color.Error = Reg.GetRegistryValue("Error", &H3232FF)
    Color.info = Reg.GetRegistryValue("Info", &HFFFF&)
    Color.BotInfo = Reg.GetRegistryValue("BotInfo", &H32FF32)
    Color.BotError = Reg.GetRegistryValue("BotError", &H3232FF)
    Color.Join = Reg.GetRegistryValue("Join", &H32FF32)
    Color.Left = Reg.GetRegistryValue("Left", &H32FF32)
    Color.Enc = Reg.GetRegistryValue("Enc", &HFF0000)
    Color.WhisperFrom = Reg.GetRegistryValue("WhisperFrom", &HFF0000)
    Color.timestamp = Reg.GetRegistryValue("TimeStamp", &H80000005)
    Color.Bot = Reg.GetRegistryValue("Bot", &H7F7F7F)
    Color.News = Reg.GetRegistryValue("News", &H32FF32)
    Color.Background = Reg.GetRegistryValue("Background", &H0&)
    
    If BNET.varShowSvr = 1 Then
        svrBNCS = "[BNCS] "
        svrBNLS = "[BNLS] "
        svrFTP = "[FTP] "
        svrMCP = "[MCP] "
        svrD2GS = "[D2GS] "
    Else
        svrBNCS = ""
        svrBNLS = ""
        svrFTP = ""
        svrMCP = ""
        svrD2GS = ""
    End If
    
    modFunctions.SaveConfig
End Function
Public Function SaveConfig()
On Local Error Resume Next
    Dim strRegLoc As String, tmpLoc As String
    tmpLoc = App.Path
    If InStr(tmpLoc, "\") Then tmpLoc = Replace(tmpLoc, "\", "/")
    strRegLoc = "Software\OpenFyre\FyreChat\" & tmpLoc
    Set Reg = New clsRegistry
    Reg.hkey = HKEY_Local_MACHINE
    Reg.KeyRoot = strRegLoc
    Reg.Subkey = "Config"
    Call Reg.SetRegistryValue("Username", BNET.username, REG_SZ)
    Call Reg.SetRegistryValue("Password", BNET.password, REG_SZ)
    Call Reg.SetRegistryValue("Server", BNET.BNCSServer, REG_SZ)
    Call Reg.SetRegistryValue("BNLS Server", BNET.BNLSServer, REG_SZ)
    Call Reg.SetRegistryValue("Product", BNET.Product, REG_SZ)
    
    ' Fix BNET Keys
    BNET.CDKey = UCase(BNET.CDKey)
    BNET.CDKey = Replace(BNET.CDKey, " ", "")
    BNET.CDKey2 = UCase(BNET.CDKey2)
    BNET.CDKey2 = Replace(BNET.CDKey2, " ", "")
    Call Reg.SetRegistryValue("CDKey", BNET.CDKey, REG_SZ)
    Call Reg.SetRegistryValue("CDKey2", BNET.CDKey2, REG_SZ)
    Call Reg.SetRegistryValue("Owner", BNET.CDKeyOwner, REG_SZ)
    Call Reg.SetRegistryValue("Channel", BNET.HomeChannel, REG_SZ)
    Call Reg.SetRegistryValue("Realm AutoLogin", BNET.varCRealm, REG_DWORD)
    Call Reg.SetRegistryValue("Character", BNET.Character, REG_SZ)
    Call Reg.SetRegistryValue("Account Email", BNET.Email, REG_SZ)
    Call Reg.SetRegistryValue("AutoConnect", BNET.varAutoCon, REG_DWORD)
    Call Reg.SetRegistryValue("Reconnect", BNET.varARCon, REG_DWORD)
    Call Reg.SetRegistryValue("LagPlug", BNET.varLagPlug, REG_SZ)
    Call Reg.SetRegistryValue("UDP", BNET.varUDP, REG_DWORD)
    Call Reg.SetRegistryValue("Join Product", BNET.varProdSpec, REG_DWORD)
    Call Reg.SetRegistryValue("Join/Leave Notifications", BNET.varEnterLeave, REG_DWORD)
    Call Reg.SetRegistryValue("Status Update", BNET.varStatusUpdate, REG_DWORD)
    Call Reg.SetRegistryValue("PingFlags", BNET.varPingFlag, REG_DWORD)
    Call Reg.SetRegistryValue("Logging", BNET.varLog, REG_DWORD)
    Call Reg.SetRegistryValue("Debug Mode", BNET.varDebugMode, REG_DWORD)
    Call Reg.SetRegistryValue("Advanced Debug Mode", BNET.varAdvDebug, REG_DWORD)
    Call Reg.SetRegistryValue("Show Server Types", BNET.varShowSvr, REG_DWORD)
    Call Reg.SetRegistryValue("Whisper Window", BNET.WhispWin, REG_DWORD)
    Call Reg.SetRegistryValue("Show News", BNET.varNews, REG_DWORD)
    Call Reg.SetRegistryValue("MinToTray", BNET.varSysTray, REG_DWORD)
    Call Reg.SetRegistryValue("Exit Prompt", BNET.varConfirmExit, REG_DWORD)
    Call Reg.SetRegistryValue("Flood Protection", BNET.FloodProt, REG_DWORD)
    Call Reg.SetRegistryValue("Max Chat Length", BNET.MaxLength, REG_SZ)
    Call Reg.SetRegistryValue("Away Idle", BNET.varAwayIdle, REG_DWORD)
    Call Reg.SetRegistryValue("Idle Interval", BNET.IdleInt, REG_SZ)
    Call Reg.SetRegistryValue("Idle Text", BNET.Idle, REG_SZ)
    Call Reg.SetRegistryValue("Normal Idle", BNET.varIdle, REG_DWORD)
    Call Reg.SetRegistryValue("Hmm Idle", BNET.varCountIdle, REG_DWORD)
    Call Reg.SetRegistryValue("Fonts", BNET.Fonts, REG_SZ)
    Call Reg.SetRegistryValue("Font Size", BNET.FSize, REG_SZ)
    Call Reg.SetRegistryValue("BackGround Color", BNET.BGColors, REG_SZ)
    
    'Reg.Subkey = "WebList"
    'Weblist
    Call Reg.SetRegistryValue("WebList Enabled", blWL_Enabled, REG_DWORD)
    Call Reg.SetRegistryValue("FTP Address", strWL_FTPAddress, REG_SZ)
    Call Reg.SetRegistryValue("FTP Location", strWL_FTPLoc, REG_SZ)
    Call Reg.SetRegistryValue("FTP User", strWL_FTPUser, REG_SZ)
    Call Reg.SetRegistryValue("FTP Password", strWL_FTPPass, REG_SZ)
    Call Reg.SetRegistryValue("FTP Upload Interval", intWL_UpTimer, REG_SZ)

    Reg.Subkey = "Colors"
    Call Reg.SetRegistryValue("Self", Color.Self, REG_SZ)
    Call Reg.SetRegistryValue("Op", Color.Op, REG_SZ)
    Call Reg.SetRegistryValue("User", Color.user, REG_SZ)
    Call Reg.SetRegistryValue("Message", Color.Message, REG_SZ)
    Call Reg.SetRegistryValue("Carrot", Color.Carrot, REG_SZ)
    Call Reg.SetRegistryValue("PingFlags", Color.PingFlags, REG_SZ)
    Call Reg.SetRegistryValue("Error", Color.Error, REG_SZ)
    Call Reg.SetRegistryValue("Info", Color.info, REG_SZ)
    Call Reg.SetRegistryValue("BotInfo", Color.BotInfo, REG_SZ)
    Call Reg.SetRegistryValue("BotError", Color.BotError, REG_SZ)
    Call Reg.SetRegistryValue("Join", Color.Join, REG_SZ)
    Call Reg.SetRegistryValue("Left", Color.Left, REG_SZ)
    Call Reg.SetRegistryValue("Enc", Color.Enc, REG_SZ)
    Call Reg.SetRegistryValue("WhisperFrom", Color.WhisperFrom, REG_SZ)
    Call Reg.SetRegistryValue("TimeStamp", Color.timestamp, REG_SZ)
    Call Reg.SetRegistryValue("Bot", Color.Bot, REG_SZ)
    Call Reg.SetRegistryValue("News", Color.News, REG_SZ)
    Call Reg.SetRegistryValue("Background", Color.Background, REG_SZ)
End Function

Public Function ToHex(data As String) As String
    Dim i As Integer
    For i = 1 To Len(data)
        ToHex = ToHex & Right("00" & Hex(Asc(Mid(data, i, 1))), 2)
    Next i
End Function

Public Function GetWORD(data As String) As Long
    Dim lReturn As Long
    Call CopyMemory(lReturn, ByVal data, 2)
    GetWORD = lReturn
End Function

Public Sub AddChatDebug(ParamArray saElements() As Variant)
    On Local Error Resume Next
    
    Dim i As Integer
    Dim data As String
    
    For i = LBound(saElements) To UBound(saElements) Step 2
        With frmMain.Chat_Output
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = saElements(i)
            .SelText = saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
            .SelStart = Len(.text)
        End With
        
        data = data & saElements(i + 1)
    Next i
    
    Call CheckMaxLength(frmMain.Chat_Output)
    
    If BNET.varLog = 1 Then Log data
End Sub

Public Sub AddChat(ParamArray saElements() As Variant)
    On Local Error Resume Next
    
    Dim i As Integer
    Dim data As String, timestamp As String
    timestamp = "[" & Format(Time, "hh:mm:ss") & "] "
    
    With frmMain.Chat_Output
        .SelStart = Len(.text)
        .SelLength = 0
        .SelColor = Color.timestamp
        .SelText = timestamp
        .SelStart = Len(.text)
        data = timestamp
    End With
    
    For i = LBound(saElements) To UBound(saElements) Step 2
        With frmMain.Chat_Output
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = saElements(i)
            .SelText = saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
            .SelStart = Len(.text)
        End With
        
        data = data & saElements(i + 1)
    Next i
    
    Call CheckMaxLength(frmMain.Chat_Output)
    
    If BNET.varLog = 1 Then Log data
End Sub

Public Sub AddWhisper(ParamArray saElements() As Variant)
On Local Error Resume Next
    Dim i As Integer, data As String, timestamp As String
    timestamp = "[" & Format(Time, "hh:mm:ss") & "] "
    If BNET.WhispWin = 1 Then
        With frmMain.rtbWhisper
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = Color.timestamp
            .SelText = timestamp
            .SelStart = Len(.text)
            data = timestamp
        End With
        For i = LBound(saElements) To UBound(saElements) Step 2
            With frmMain.rtbWhisper
                .SelStart = Len(.text)
                .SelLength = 0
                .SelColor = saElements(i)
                .SelText = saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
                .SelStart = Len(.text)
            End With
            data = data & saElements(i + 1)
        Next i
        Call CheckMaxLength(frmMain.rtbWhisper)
    Else
        With frmMain.Chat_Output
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = Color.timestamp
            .SelText = timestamp
            .SelStart = Len(.text)
            data = timestamp
        End With
        For i = LBound(saElements) To UBound(saElements) Step 2
            With frmMain.Chat_Output
                .SelStart = Len(.text)
                .SelLength = 0
                .SelColor = saElements(i)
                .SelText = saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
                .SelStart = Len(.text)
            End With
            data = data & saElements(i + 1)
        Next i
        Call CheckMaxLength(frmMain.Chat_Output)
    End If
    If BNET.varLog = 1 Then Log data
End Sub

Public Sub Send(ByVal Message As String, SOCKET As Winsock, Optional Extra As Boolean, Optional ByVal flags As Long)
On Local Error Resume Next
    If Not blEnterChat Then Exit Sub
    Dim WaitTime As Long
    Dim Queue As ListBox
    If Message = "" Then Message = " "
    Dim strData As String
    WaitTime = RequiredDelay(Len(Message))
    Set Queue = frmMain.lstEnqueue
    lSendDelay = WaitTime
    If BNET.FloodProt Then
        If (lSendDelay <> 0) Or (frmMain.lstEnqueue.ListCount > 0) Then
            strData = strData & Message
        ElseIf frmMain.lstEnqueue.ListCount = 0 Then
            With sPB
                .InsertNTString Message
                .SendPacket &HE
            End With
            Call PrintUser(Message)
            If (Len(Message) >= 1) Then
                If awayFlag Then
                    If InStr(Message, "/away Bot has been idle for") Then
                        tFlag = True
                        Exit Sub
                    End If
                    If (Mid(Message, 1, 1) = "/") Then
                        If (Mid(Message, 1, 2) = "/w") Or (Mid(Message, 1, 2) = "/m") Then tFlag = False
                        If (Mid(Message, 1, 6) = "/whois") Or (Mid(Message, 1, 7) = "/whoami") Then tFlag = True
                    Else
                        tFlag = False
                    End If
                    If tFlag Then
                        Exit Sub
                    Else
                        Send "/away", frmMain.wsBNET
                        tFlag = False
                        LastTalk = GetTickCount()
                    End If
                End If
            End If
            Exit Sub
        End If
        Queue.AddItem strData
    Else
        sPB.InsertNTString Message
        sPB.SendPacket &HE
        Call PrintUser(Message)
        If (Len(Message) >= 1) Then
            If awayFlag Then
                If InStr(Message, "/away Bot has been idle for") Then
                    tFlag = True
                    Exit Sub
                End If
                If (Mid(Message, 1, 1) = "/") Then
                    If (Mid(Message, 1, 2) = "/w") Or (Mid(Message, 1, 2) = "/m") Then tFlag = False
                    If (Mid(Message, 1, 6) = "/whois") Or (Mid(Message, 1, 6) = "/whoami") Then tFlag = True
                Else
                    tFlag = False
                End If
                If tFlag Then
                    Exit Sub
                Else
                    Send "/away", frmMain.wsBNET
                    awayFlag = False
                    tFlag = False
                    LastTalk = GetTickCount()
                End If
            End If
        End If
        Exit Sub
    End If
End Sub

Public Sub Log(data As String)
On Local Error Resume Next
    Dim fHandle As Long
    fHandle = FreeFile()
    Call IsPathAFolder(App.Path & "/logs")
    If PathFolder Then
        Open (App.Path & "/logs/" & Format(Now, "mm-dd-yyyy") & ".txt") For Append As #fHandle
    Else
        Dim blah As SECURITY_ATTRIBUTES
        blah.nLength = LenB(blah)
        Call CreateDirectory(App.Path & "/logs", blah)
        Open (App.Path & "/logs/" & Format(Now, "mm-dd-yyyy") & ".txt") For Append As #fHandle
    End If
    Print #fHandle, data
    Close #fHandle
End Sub

Public Function IsPathAFolder(ByVal sPath As String) As Boolean
On Local Error Resume Next
    Dim Result As Long
    Result = PathIsDirectory(sPath)
    IsPathAFolder = (Result = vbDirectory) Or (Result = 1)
    If Result = 16 Then
        PathFolder = True
    Else
        PathFolder = False
    End If
End Function

Public Function WildCard(ByVal data As String, lstWild As ListView) As String
On Local Error Resume Next
    Dim OpSplit() As String, Message As String
    Dim cardMode As WildCard
    Dim i As Integer
    
    OpSplit = Split(data, " ")
    cardMode = WildCardMode(OpSplit(0))
    Message = ""
    
    If InStr(data, " ") Then
        Message = (Mid(data, InStr(data, " ") + 1))
        data = Mid(data, 1, InStr(data, " ") - 1)
    End If
    
    If InStr(data, "*") Then data = Replace(data, "*", "")

    With lstWild
        Select Case cardMode
            Case 1
                For i = 1 To .ListItems.count
                    If InStr(LCase(.ListItems(i).text), LCase(data)) Then
                        If LCase(Mid(data, 1, Len(.ListItems(i).text))) = LCase(data) Then
                            If Message = "" Then
                                WildCard = WildCard & ", " & .ListItems(i).text
                                'WildCard = .ListItems(i).text
                            Else
                                WildCard = .ListItems(i).text & " " & Message
                            End If
                        End If
                    End If
                Next i
            Case 2
                For i = 1 To .ListItems.count
                    If LCase(Mid(data, 1, Len(.ListItems(i).text))) = LCase(data) Then
                        If LCase(StrReverse(Mid(data, 1, Len(.ListItems(i).text)))) = LCase(StrReverse(data)) Then
                            If Message = "" Then
                                WildCard = WildCard & ", " & .ListItems(i).text
                                'WildCard = .ListItems(i).text
                            Else
                                WildCard = .ListItems(i).text & " " & Message
                            End If
                        End If
                    End If
                Next i
            Case 3
                For i = 1 To .ListItems.count
                    If InStr(LCase(.ListItems(i).text), LCase(data)) Then
                        If Message = "" Then
                            WildCard = WildCard & ", " & .ListItems(i).text
                            'WildCard = .ListItems(i).text
                        Else
                            WildCard = .ListItems(i).text & " " & Message
                        End If
                    End If
                Next i
            Case Else
                If Message = "" Then
                    WildCard = data
                Else
                    WildCard = data & " " & Message
                End If
        End Select
    End With
End Function

Public Function DisplayFound(data As String, distype As String) As String
On Local Error Resume Next
    Dim user As String, users() As String
    Dim cont As Boolean
    Dim i As Integer, J As Integer
    data = Mid(Replace(data, ",", ""), 2)
    user = data
    If InStr(data, " ") Then
        users = Split(data, " ")
        user = users(0)
        cont = True
        J = 0
    Else
        J = 1
        cont = False
    End If
    With frmMain.lstChannel
        For i = 1 To .ListItems.count
            If (LCase(.ListItems(i)) = LCase(user)) Then
                Select Case distype
                    Case "getlag"
                        DisplayFound = DisplayFound & user & " => " & .FindItem(user).ListSubItems(1).ToolTipText & vbNewLine
                        If (cont) Then
                            J = J + 1
                            user = users(J)
                        Else
                            Exit For
                        End If
                    Case "getinfo"
                        DisplayFound = DisplayFound & user & " => " & .FindItem(user).Tag & vbNewLine
                        If (cont) Then
                            J = J + 1
                            user = users(J)
                        Else
                            Exit For
                        End If
                    Case Else
                        DisplayFound = "null"
                End Select
            End If
        Next i
    End With
    Counter = J
    DisplayFound = Left(DisplayFound, Len(DisplayFound) - 2)
End Function

Public Function DoAddToSendList(text As String)
    frmMain.User_Input.AddItem text, 0
    If frmMain.User_Input.ListCount > 10 Then frmMain.User_Input.RemoveItem 10
End Function

Public Function ClearBuffers()
    frmMain.Chat_Output.text = ""
    frmMain.rtbWhisper.text = ""
    frmMain.User_Input.Clear
End Function

Public Function DecodeBase8(mystring As String) As String
On Local Error Resume Next
    Dim d As Integer, i As Integer, c As Integer, charSet As String
    charSet = "[]<>.,*?"
    
    For d = 1 To Len(mystring)
        c = c + (InStr(1, charSet, Mid(mystring, d, 1)) - 1) * (8 ^ i)
        i = i + 1
        If i = 3 Then
            DecodeBase8 = DecodeBase8 & Chr$(c)
            i = 0
            c = 0
        End If
    Next
End Function

Public Function CVL(x As String) As Long
On Local Error Resume Next
    If Len(x) < 4 Then
        MsgBox "CVL(): String too short"
        Exit Function
    End If
    CopyMemory CVL, ByVal x, 4
End Function

Public Function MakeServer(data As String) As String
    MakeServer = CLng("&H" & ToHex(Mid(data, 1, 1))) & "." & CLng("&H" & ToHex(Mid(data, 2, 1))) & "." & CLng("&H" & ToHex(Mid(data, 3, 1))) & "." & CLng("&H" & ToHex(Mid(data, 4, 1)))
End Function

Public Function FormatCount(count As Long, Optional FormatType As TimeFormatType = 4) As String
On Local Error Resume Next
    Dim Days As Long, Hours As Long, Minutes As Long, Seconds As Long, Miliseconds As Long
    Miliseconds = count Mod 1000
    count = count \ 1000
    Days = count \ (24& * 3600&)
    If Days > 0 Then count = count - (24& * 3600& * Days)
    Hours = count \ 3600&
    If Hours > 0 Then count = count - (3600& * Hours)
    Minutes = count \ 60
    Seconds = count Mod 60
    Select Case FormatType
        Case 0
            FormatCount = Days & " days, " & Hours & " hours, " & _
            Minutes & " minutes, " & Seconds & " seconds, and " & Miliseconds & _
            " milliseconds"
        Case 1
            FormatCount = Days & " days, " & Hours & " hours, " & _
            Minutes & " minutes, and " & Seconds & " seconds"
        Case 2
            FormatCount = Days & ":" & Hours & ":" & _
            Minutes & ":" & Seconds & ":" & Miliseconds
        Case 3
            Select Case Minutes
                Case "0"
                Case "1"
                    FormatCount = Minutes & " minute"
                Case Else
                    FormatCount = Minutes & " minutes"
            End Select
            Select Case Hours
                Case "0"
                Case "1"
                    FormatCount = Hours & " hour and " & Minutes & " minutes"
                Case Else
                    FormatCount = Hours & " hours and " & Minutes & " minutes"
            End Select
            Select Case Days
                Case "0"
                Case "1"
                    FormatCount = Days & " day, " & Hours & " hours and " & Minutes & " minutes"
                Case Else
                    FormatCount = Days & " days, " & Hours & " hours and " & Minutes & " minutes"
            End Select
        Case 4
            Select Case Seconds
                Case "0"
                    FormatCount = "0." & Format(Miliseconds, "0000")
                Case Else
                    FormatCount = Seconds & " seconds"
            End Select
            Select Case Minutes
                Case "0"
                Case "1"
                    FormatCount = Minutes & " minute, and " & Seconds & " seconds"
                Case Else
                    FormatCount = Minutes & " minutes, and " & Seconds & " seconds"
            End Select
            Select Case Hours
                Case "0"
                Case "1"
                    FormatCount = Hours & " hour, " & Minutes & " minutes, and " & Seconds & " seconds"
                Case Else
                    FormatCount = Hours & " hours, " & Minutes & " minutes, and " & Seconds & " seconds"
            End Select
            Select Case Days
                Case "0"
                Case "1"
                    FormatCount = Days & " day, " & Hours & " hours, " & Minutes & " minutes, and " & Seconds & " seconds"
                Case Else
                    FormatCount = Days & " days, " & Hours & " hours, " & Minutes & " minutes, ands " & Seconds & " seconds"
            End Select
        End Select
End Function

Public Sub RequestProfile(user As String)
On Local Error Resume Next
    Dim lngKey As Long
    lngKey = GetTickCount()
    varRequest = "Profile"
    ProfUser = user
    If Mid(ProfUser, 1, 1) = "," Then ProfUser = Mid(ProfUser, 3)
    With sPB
        .InsertDWORD 1           'Number of accounts
        .InsertDWORD 4           'Number of keys
        .InsertDWORD lngKey      'Request ID
        .InsertNTString ProfUser 'Requested Accounts
        'Requested Keys
        .InsertNTString "profile\sex"
        .InsertNTString "profile\age"
        .InsertNTString "profile\location"
        .InsertNTString "profile\description"
        .SendPacket &H26
    End With
End Sub

Public Function RequestRecordData(user As String)
On Local Error Resume Next
    If InStr(user, "#") Then user = BNET.username
    Dim lngKey As Long
    lngKey = GetTickCount()
    varRequest = "RecordData"

    With sPB
        .InsertDWORD 1          'Number of accounts
        .InsertDWORD 6          'Number of keys
        .InsertDWORD lngKey     'Request ID
        .InsertNTString user    'Requested Accounts
        'Requested Keys
        .InsertNTString "System\Account Created"
        .InsertNTString "System\Username"
        .InsertNTString "System\Last Logon"
        .InsertNTString "System\Account Expires"
        .InsertNTString "System\Last Logoff"
        .InsertNTString "System\Time Logged"
        .SendPacket &H26
    End With
End Function

Public Function GetDWORD(data As String) As Long
    Dim lReturn As Long
    Call CopyMemory(lReturn, ByVal data, 4)
    GetDWORD = lReturn
End Function

Public Sub AddToTray(TrayIcon, TrayText As String, TrayForm As Form)
On Local Error Resume Next
    nid.cbSize = Len(nid)
    nid.hWnd = TrayForm.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = TrayIcon
    nid.szTip = TrayText & vbNullChar
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Public Sub ModifyTray(TrayIcon, TrayText As String, TrayForm As Form)
On Local Error Resume Next
    nid.cbSize = Len(nid)
    nid.hWnd = TrayForm.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = TrayIcon
    nid.szTip = TrayText & vbNullChar
    Shell_NotifyIcon NIM_MODIFY, nid
End Sub

Public Sub RemoveFromTray(): Shell_NotifyIcon NIM_DELETE, nid: End Sub

Public Function RespondToTray(x As Single)
On Local Error Resume Next
    Dim AcessMenuTray As Boolean
    RespondToTray = 0
    Dim Msg As Long
    If frmMain.ScaleMode <> 3 Then Msg = x / Screen.TwipsPerPixelX Else: Msg = x
    Select Case Msg
        Case WM_LBUTTONDOWN
        Case WM_LBUTTONUP
        Case WM_LBUTTONDBLCLK
            RespondToTray = 1
        Case WM_RBUTTONDOWN
            Call frmMain.TrayPop
        Case WM_RBUTTONUP
        Case WM_RBUTTONDBLCLK
    End Select
End Function

Public Sub ShowFormAgain()
On Local Error Resume Next
    If vbMax Then
        frmMain.WindowState = 2
    Else
        frmMain.WindowState = 0
    End If
    frmMain.Show
    App.TaskVisible = True
 End Sub
 
Public Function hexEncrypt(ByVal SString As String) As String
On Local Error Resume Next
    Dim sHex As String
    Dim i As Long
    Dim pos As Long
    Dim Encrypt As Boolean
    Dim sNew As String
    Dim sTmp As String
    Dim iDec As Long
    pos = 1
    For i = 1 To Len(SString)
    If Mid(SString, 1, 1) <> Chr(163) Then
        sHex = sHex & Hex$(Asc(Mid(SString, i, 1)))
        If Len(sHex) = 1 Then sHex = "0" & sHex
            Encrypt = True
        Else
            sTmp = Mid(SString, 2, Len(SString))
            sHex = Mid(sTmp, pos, 2)
            iDec = Val("&H" & sHex)
            If iDec > 0 Then
                sNew = sNew & Chr(iDec)
            End If
            pos = pos + 2
            Encrypt = False
        End If
        Next
        If Encrypt Then
            hexEncrypt = Chr(163) & sHex
        Else
            hexEncrypt = sNew
        End If
End Function

Public Function DownloadFile(location As String, FileSave As String) As Boolean
    DownloadFile = URLDownloadToFile(0&, location, FileSave, 0&, 0&) = ERROR_SUCCESS
End Function

Public Sub Wait(ByVal Duration As Double)
On Local Error Resume Next
    Dim StartTime As Double
    StartTime = Timer
    Do While Timer - StartTime < Duration
        DoEvents
    Loop
End Sub

Public Function WildCardMode(ByVal InString As String) As WildCard
On Local Error Resume Next
    If (Left(InString, 1) = "*") And (Right(InString, 1) = "*") Then
        WildCardMode = AsteriskOnBoth
        Exit Function
    ElseIf (Left(InString, 1) = "*") Then
        WildCardMode = AsteriskOnLeft
        Exit Function
    ElseIf (Right(InString, 1) = "*") Then
        WildCardMode = AsteriskOnRight
        Exit Function
    Else
        WildCardMode = None
        Exit Function
    End If
End Function

Public Function RequiredDelay(ByVal Bytes As Long) As Long
On Local Error Resume Next
    Static LastTick As Long
    Static SentBytes As Long
    Const perPacket = 250
    Const perByte = 15
    Const MaxBytes = 400
    
    Dim Tick As Long
    Tick = GetTickCount()
    
    If Abs(LastTick - Tick) > (SentBytes * perByte) Then
    SentBytes = 0
    Else
    SentBytes = SentBytes - (Abs(LastTick - Tick) / perByte)
    End If
    
    LastTick = Tick
    If (SentBytes + perPacket + Bytes) > MaxBytes Then
        If (Bytes > 200) Or ((Bytes > 25) And ((SentBytes - perPacket) > 400)) Then
            RequiredDelay = (((SentBytes + perPacket + Bytes) - MaxBytes) * perByte) + _
            (1500 * (((SentBytes + Bytes) - ((SentBytes + Bytes) Mod 200)) / 200))
            If Bytes > 200 Then
                RequiredDelay = RequiredDelay + Bytes + (((Bytes - 200) - ((Bytes - 200) Mod 10)) * 50)
            End If
            SentBytes = SentBytes + Bytes
        Else
            RequiredDelay = ((SentBytes + perPacket + Bytes) - MaxBytes) * perByte
        End If
    Else
        SentBytes = SentBytes + perPacket + Bytes
        RequiredDelay = 0
    End If
End Function

Public Sub PrintUser(Message As String)
On Local Error Resume Next
    If Left(Message, 1) = "/" Then Exit Sub
    Dim icon As Integer
    
    Dim eData, mData As String
    Select Case Mid$(Message, 1, 1)
        Case Chr(166)
            eData = " (Octal)"
            mData = DecodeBase8(Mid$(Message, 2, Len(Message)))
        Case "£"
            eData = " (Hex)"
            mData = hexEncrypt(Message)
        Case "·"
            eData = " (Reversed)"
            mData = StrReverse(Mid(Message, 2))
        Case Else
            eData = ""
            mData = Message
    End Select
    
    PingFlags = ""
    If BNET.varPingFlag = 1 Then PingFlags = "[" & MyPing & ":" & MyFlags & "] "
    AddChat Color.PingFlags, PingFlags, Color.Self, "<" & BNET.TrueUsername, Color.Enc, _
        eData, Color.Self, "> ", Color.Message, mData
End Sub

Public Function CheckMaxLength(ByRef Chat_Output As RichTextBox) As Long
On Local Error Resume Next
    Static Checking As Boolean
    Dim i As Integer
    If Checking = True Then Exit Function
    Checking = True
    Do While (Len(Chat_Output.text) > BNET.MaxLength)
        With Chat_Output
            i = InStr(1, Chat_Output.text, vbCrLf)
            If i Then i = i + 1
            .Visible = False
            .SelStart = 0
            .SelLength = i
            CheckMaxLength = CheckMaxLength + i
            .SelText = ""
            .SelStart = Len(.TextRTF)
            .SelLength = 0
            .Visible = True
        End With
    Loop
    Checking = False
End Function

Public Function FileExists(sFileName As String) As Boolean
    If Len(sFileName$) = 0 Then
        FileExists = False
        Exit Function
    End If
    If Len(Dir$(sFileName$)) Then
        FileExists = True
    Else
        FileExists = False
    End If
End Function

Public Function MakeLong(x As String) As Long
    If Len(x) < 4 Then Exit Function
    CopyMemory MakeLong, ByVal x, 4
End Function

Public Sub strcpy(ByRef Source As String, ByVal nText As String): Source = Source & nText: End Sub

Public Function KillNull(ByVal text As String) As String
    Dim i As Integer
    i = InStr(1, text, Chr(0))
    If i = 0 Then
        KillNull = text
        Exit Function
    End If
    KillNull = Left(text, i - 1)
End Function

Public Function LShift(ByVal pnValue As Long, ByVal pnShift As Long) As Double
    LShift = CDbl(pnValue * (2 ^ pnShift))
End Function

Public Function RShift(ByVal pnValue As Long, ByVal pnShift As Long) As Double
    RShift = CDbl(pnValue \ (2 ^ pnShift))
End Function

Public Function HexToStr(ByVal Hex1 As String) As String
    Dim strTemp As String, strReturn As String, i As Long
    If Len(Hex1) Mod 2 <> 0 Then Exit Function
    For i = 1 To Len(Hex1) Step 2
        strReturn = strReturn & Chr(Val("&H" & Mid(Hex1, i, 2)))
    Next i
    HexToStr = strReturn
End Function

Public Function StrToHex(ByVal String1 As String) As String
    Dim strTemp As String, strReturn As String, i As Long
    For i = 1 To Len(String1)
        strTemp = Hex(Asc(Mid(String1, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & " " & strTemp
    Next i
    StrToHex = strReturn
End Function

Public Sub NullTruncString(ByRef text As String)
    Dim i As Integer
    i = InStr(text, Chr(0))
    If i = 0 Then Exit Sub
    text = Left(text, i - 1)
End Sub

'By Grok[vL]
Public Function DebugOutput(ByVal sIn As String) As String
On Local Error Resume Next
    Dim x1 As Long, y1 As Long
    Dim iLen As Long, iPos As Long
    Dim sB As String, sT As String
    Dim sOut As String
    Dim Offset As Long, sOffset As String

    iLen = Len(sIn)
    If iLen = 0 Then Exit Function
    sOut = ""
    Offset = 0
    For x1 = 0 To ((iLen - 1) \ 16)
        sOffset = Right$("0000" & Hex(Offset), 4)
        sB = String(48, " ")
        sT = "................"
    For y1 = 1 To 16
        iPos = 16 * x1 + y1
    If iPos > iLen Then Exit For
    Mid(sB, 3 * (y1 - 1) + 1, 2) = Right("00" & Hex(Asc(Mid(sIn, iPos, 1))), 2) & " "
    Select Case Asc(Mid(sIn, iPos, 1))
        Case 0, 9, 10, 13
        Case Else
            Mid(sT, y1, 1) = Mid(sIn, iPos, 1)
    End Select
    Next y1
    If Len(sOut) > 0 Then sOut = sOut & vbCrLf
    sOut = sOut & sOffset & ": "
    sOut = sOut & sB & " " & sT
    Offset = Offset + 16
    Next x1
    DebugOutput = sOut
End Function

Public Sub DisplayAdvDebug(s As String, d As String, p As Byte)
    Select Case s
        Case "BNCS"
            If Not ((p = &H25 Or p = &H0 Or p = &HF) And blEnterChat = True) Then
                AddChat Color.BotInfo, "[" & s & "] ", Color.Carrot, "[S 0x" & Hex(p) & "] ", Color.BotInfo, PNBNCS(p)
                AddChatDebug Color.Bot, DebugOutput(d)
            End If
        Case "BNLS"
            AddChat Color.BotInfo, "[" & s & "] ", Color.Carrot, "[S 0x" & Hex(p) & "] ", Color.BotInfo, PNBNLS(p)
            AddChatDebug Color.Bot, DebugOutput(d)
        Case "MCP"
            AddChat Color.BotInfo, "[" & s & "] ", Color.Carrot, "[S 0x" & Hex(p) & "] ", Color.BotInfo, PNMCP(p)
            AddChatDebug Color.Bot, DebugOutput(d)
        Case "D2GS"
            AddChat Color.BotInfo, "[" & s & "] ", Color.Carrot, "[S 0x" & Hex(p) & "] ", Color.BotInfo, PNSD2GS(p)
            AddChatDebug Color.Bot, DebugOutput(d)
        Case Else
            AddChat Color.BotInfo, "[" & s & "] ", Color.Carrot, "[S 0x" & Hex(p) & "] ", Color.BotInfo, "Unknown"
            AddChatDebug Color.Bot, DebugOutput(d)
    End Select
End Sub

Public Function ProductID(Client As String) As String
    Select Case LCase(Client)
        Case "starcraft": ProductID = "RATS"
        Case "starcraft japan": ProductID = "RTSJ"
        Case "starcraft brood war": ProductID = "PXES"
        Case "starcraft shareware": ProductID = "RHSS"
        Case "diablo": ProductID = "LTRD"
        Case "diablo shareware": ProductID = "RHSD"
        Case "diablo ii": ProductID = "VD2D"
        Case "diablo ii lord of destruction": ProductID = "PX2D"
        Case "warcraft ii battle.net edition": ProductID = "NB2W"
        Case "warcraft iii": ProductID = "3RAW"
        Case "warcraft iii the frozen throne": ProductID = "PX3W"
        Case Else: ProductID = "RATS"
    End Select
End Function

Public Function ProductName(Client As String)
    Select Case Client
        Case "RATS": ProductName = "StarCraft"
        Case "RTSJ": ProductName = "StarCraft Japan"
        Case "PXES": ProductName = "StarCraft Brood War"
        Case "RHSS": ProductName = "StarCraft Shareware"
        Case "LTRD": ProductName = "Diablo"
        Case "RHSD": ProductName = "Diablo Shareware"
        Case "VD2D": ProductName = "Diablo II"
        Case "PX2D": ProductName = "Diablo II Lord of Destruction"
        Case "NB2W": ProductName = "Warcraft II Battle.net Edition"
        Case "3RAW": ProductName = "Warcraft III"
        Case "PX3W": ProductName = "Warcraft III The Frozen Throne"
        Case Else: ProductName = "Starcraft"
    End Select
End Function

Public Function GetFTSystem() As FILETIME
Dim sysTime As SYSTEMTIME
Dim fTime As FILETIME
  On Local Error Resume Next
  GetSystemTime sysTime
  SystemTimeToFileTime sysTime, fTime
  GetFTSystem = fTime
End Function

Public Function GetFTLocal() As FILETIME
Dim sysTime As SYSTEMTIME
Dim fTime  As FILETIME
  On Local Error Resume Next
  GetLocalTime sysTime
  SystemTimeToFileTime sysTime, fTime
  GetFTLocal = fTime
End Function

Public Sub BNETFTP(ByVal FileName As String, Optional ByVal LFTimeLow As Long = 0, Optional ByVal LFTimeHigh As Long = 0, Optional ByVal StartPos As Long = 0, Optional ByVal bID As Long = 0, Optional ByVal bExt As Long = 0)
    frmMain.wsDL.SendData Chr(2)
    With sPB
        .InsertByte 0                               ' (WORD) Request Length
        .InsertByte 1                               ' (WORD) Protocol Version)
        .InsertNonNTString "68XI" & BNET.Product    ' (DWORD) Platform ID & (DWORD) Product ID
        .InsertDWORD bID                            ' (DWORD) Banner ID
        .InsertDWORD bExt                           ' (DWORD) Banner File Extension
        .InsertDWORD StartPos                       ' (DWORD) Start position in file (for resuming)
        If BnetFileTime = "" Then
            .InsertDWORD &H0
            .InsertDWORD &H0
        Else
            .InsertDWORD CLng(Split(BnetFileTime, " ")(1))
            .InsertDWORD CLng(Split(BnetFileTime, " ")(0))
            '.InsertDWORD LFTimeLow                      ' (DWORD) Filetime of local file
            '.InsertDWORD LFTimeHigh                     '   cont. FileTime
        End If
        .InsertNTString FileName                    ' (DWORD) Filename
        .SendServer frmMain.wsDL
    End With
End Sub

Public Sub ConnectBNETFTP(ByVal fName As String, Optional ByVal fTime As String = "")
    BnetFileName = fName
    BnetFileTime = fTime
    If BNET.varDebugMode = 1 Then AddChat Color.BotInfo, svrFTP, Color.Bot, "Connecting to Battle.net FTP..."
    frmMain.wsDL.Close
    frmMain.wsDL.Connect BNET.BNCSServer, 6112
End Sub
