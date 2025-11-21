Attribute VB_Name = "modMCP"
Option Explicit
Private mcpPB As New clsBuffer
Private spID As String
Private ReqID As Long
Private d2gName As String
Private d2gPass As String
Public d2char As String
Public d2charClass As String

Public Sub ParseMCP(d As String)
Dim msgLen As Long, pID As Byte, Msg As String
    With mcpPB
        .Clear
        .SetBuffer d
    
        msgLen = .GetWORD           '(WORD) Message length, including this header
        pID = .GetByte              '(BYTE) Message ID
        Msg = Mid(.GetBuffer, 4)    '(VOID) Message Data
    End With
    
    If BNET.varAdvDebug = 1 Then DisplayAdvDebug "MCP", d, pID
    
    If (CallPluginMessageNotifications(1, pID, Msg, Len(Msg))) Then
        SleepEx 0, 1
        Exit Sub
    End If
    
    spID = ""
    If BNET.varDebugMode = 1 Then: spID = "[S 0x" & Hex(pID) & "] "
    
    Select Case pID
        Case &H1: rcvStartup
        Case &H2: rcvCharCreate
        Case &H3: rcvCreateGame
        Case &H4: rcvJoinGame
        Case &H5: rcvGameList
        Case &H6: rcvGameInfo
        Case &H7: rcvCharLogon
        Case &HA: rcvCharDelete
        Case &H11: rcvReqLadData
        Case &H12: rcvMOTD
        Case &H14: rcvCreateQ
        Case &H17: rcvCharList
        Case &H18: rcvCharUpgrade
        Case &H19: rcvCharList2
        Case Else
            If BNET.varDebugMode Then
                AddChat Color.Carrot, "[Debug] Unhandled Packet ", Color.Bot, spID, Color.Message, DebugOutput(Mid(d, 5))
            End If
    End Select
End Sub

' MCP_STARTUP (0x01)
Private Sub rcvStartup()
    Select Case mcpPB.GetDWORD ' (DWORD) Result
        Case &H2, &HA - &HD
            AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.Error, "No Battle.net connection detected!"
            frmRealm.RealmDisconnect
        Case &HE
            AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.Error, "CDKey banned from realm play."
             frmRealm.RealmDisconnect
        Case &H7F
            AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.Error, _
                "Your connection has been temporarily restricted from this realm. Please try to log in at another time."
            frmRealm.RealmDisconnect
        Case Else
            sPB.InsertDWORD &H8
            sPB.SendRPacket &H19
    End Select
End Sub

' MCP_CHARCREATE (0x02)
Private Sub rcvCharCreate()
Dim icon As Integer
    Select Case mcpPB.GetDWORD ' (DWORD) Result
        Case &H0
            With frmRealm
                AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.BotInfo, _
                    "Created " & .lstClass.text & ": " & .txtCreateChar.text
                .lstCharacter.ListItems.Add , , .txtCreateChar.text
                .lstCharacter.ListItems(.lstCharacter.ListItems.Count).ListSubItems.Add 1, , .lstClass.text, icon
                .lstCharacter.ListItems(.lstCharacter.ListItems.Count).ListSubItems.Add 2, , "1", icon
                If .chkHardChar.Value = 1 Then
                    .lstCharacter.ListItems(.lstCharacter.ListItems.Count).ListSubItems.Add 3, , "X", icon
                Else
                    .lstCharacter.ListItems(.lstCharacter.ListItems.Count).ListSubItems.Add 3, , "", icon
                End If
                If .chkExpChar.Value = 1 Then
                    .lstCharacter.ListItems(.lstCharacter.ListItems.Count).ListSubItems.Add 4, , "X", icon
                Else
                    .lstCharacter.ListItems(.lstCharacter.ListItems.Count).ListSubItems.Add 4, , "", icon
                End If
            End With
        Case &H14
            AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.Error, _
                "Character already exists or maximum number of characters reached."
        Case &H15
            AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.Error, "Invalid name."
    End Select
End Sub

' MCP_CREATEGAME (0x03)
Private Sub rcvCreateGame()
Dim errMsg As String
'Dim reqID As Long, gameTok As Long, Unknown As Long
    With mcpPB
        'reqID = .GetWORD        ' (WORD) Request Id
        'gameTok = .GetWORD      ' (WORD) Game token
        'Unknown = .GetWORD      ' (WORD) Unknown - 0
        .Skip 6
    End With
    Select Case mcpPB.GetDWORD  ' (DWORD) Result
        Case &H0
            AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.BotInfo, "Game creation succeeded! Joining Game..."
            With sPB
                .InsertWORD Val("&H" & ReqID)   ' (WORD) Request ID
                .InsertNTString d2gName         ' (STRING) Game name
                .InsertNTString d2gPass         ' (STRING) Game Password
                .SendRPacket &H4                ' MCP_JOINGAME
            End With
        Case &H1E: errMsg = "Invalid game name.": GoTo Failed
        Case &H1F: errMsg = "Game already exists.": GoTo Failed
        Case &H20: errMsg = "Game servers are down.": GoTo Failed
        Case &H6E: errMsg = "A dead hardcore character cannot create games.": GoTo Failed
    End Select
    Exit Sub
Failed:
AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.Error, errMsg
End Sub

' MCP_JOINGAME (0x04)
Private Sub rcvJoinGame()
'Dim reqID As Long, gameTok As Long, Unknown As Long
Dim ipD2GS As String, gameHash As Long, errMsg As String
    With mcpPB
        'reqID = .GetWORD               ' (WORD) Request Id
        'gameTok = .GetWORD             ' (WORD) Game token
        'Unknown = .GetWORD             ' (WORD) Unknown - 0
        .Skip 6
        ipD2GS = MakeServer(.GetRaw(4)) ' (DWORD) IP of D2GS Server
        gameHash = .GetDWORD            ' (DWORD) Game hash
    End With
    Select Case mcpPB.GetDWORD          ' (DWORD) Result
        Case &H0
            AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.BotInfo, "Game joining succeeded! Closing connection to MCP..."
            'In this case, Diablo 2 terminates the connection with the MCP and initiates the connection with the D2GS.
            Call frmRealm.CloseRealm
            frmD2DV.Hide
            
            AddChat Color.BotInfo, svrD2GS, Color.Bot, "Connecting to D2GS server ", Color.Message, ipD2GS, Color.Bot, "..."
            frmMain.wsD2GS.Close
            frmMain.wsD2GS.Connect ipD2GS, 4000
        Case &H29: errMsg = "Password incorrect.": GoTo Failed
        Case &H2A: errMsg = "Game does not exist.": GoTo Failed
        Case &H2B: errMsg = "Game is full.": GoTo Failed
        Case &H2C: errMsg = "You do not meet the level requirements for this game.": GoTo Failed
        Case &H6E: errMsg = "A dead hardcore character cannot join a game.": GoTo Failed
        Case &H71: errMsg = "A non-hardcore character cannot join a game created by a Hardcore character.": GoTo Failed
        Case &H73: errMsg = "Unable to join a Nightmare game.": GoTo Failed
        Case &H74: errMsg = "Unable to join a Hell game.": GoTo Failed
        Case &H78: errMsg = "A non-expansion character cannot join a game created by an Expansion character.": GoTo Failed
        Case &H79: errMsg = "A Expansion character cannot join a game created by a non-expansion character.": GoTo Failed
        Case &H7D: errMsg = "A non-ladder character cannot join a game created by a Ladder character.": GoTo Failed
    End Select
    Exit Sub
Failed:
     AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.Error, errMsg
End Sub

' MCP_GAMELIST (0x05)
Private Sub rcvGameList()
'Dim reqID As Long, index as long
Dim numPlayers As String, status As Long, gameName As String, gameDesc As String
Dim strStatus As String
    With mcpPB
        'reqID = .GetWORD        ' (WORD) Request Id
        'index = .GetDWORD       ' (DWORD) Index
        .Skip 6
        numPlayers = .GetByte   ' (BYTE) Number of players in game
        status = .GetDWORD      ' (DWORD) Status
        gameName = .GetString   ' (STRING) Game name
        gameDesc = .GetString   ' (STRING) Game description
    End With
    If gameName = "" Then Exit Sub
    Select Case status
        Case &H300004: strStatus = "Game is available to join."
        Case &HFFFFFFFF: strStatus = "Server is down"
        Case Else: strStatus = "Unknown"
    End Select
    If Not gameDesc = "" Then gameDesc = " (" & gameDesc & ") "
    AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.BotInfo, "Game: ", Color.Message, gameName & gameDesc, _
        Color.BotInfo, " - Players: ", Color.Message, numPlayers, Color.BotInfo, " - Status: ", Color.Message, strStatus
End Sub

' MCP_GAMEINFO (0x06)
Private Sub rcvGameInfo()
Dim status As Long, gameUptime As Long, maxPlayers As Long, numChars As Long
Dim classes() As Long, levels() As Long, charNames As String
Dim i As Long
    With mcpPB
        .Skip 2                 ' (WORD) Request ID
        status = .GetDWORD      ' (DWORD) Status
        gameUptime = .GetDWORD  ' (DWORD) Game Uptime (seconds)
        .Skip 2                 ' (WORD) Unknown
        maxPlayers = .GetByte   ' (BYTE) Maximum players allowed
        numChars = .GetByte     ' (BYTE) Number of characters in the game
        For i = 0 To 16
            classes(i) = .GetByte   ' (BYTE[16]) Classes of ingame characters (Last 8 empty)
        Next i
        For i = 0 To 16
            levels(i) = .GetByte    ' (BYTE[16]) Levels of ingame characters  (Last 8 empty)
        Next i
        .Skip 1                  ' (BYTE) Unused (0)
        charNames = .GetString   ' (STRINGS) Character names

    End With
    Select Case status
        Case &H300004
            AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.BotInfo, "Uptime: ", Color.Message, FormatCount(gameUptime, DaysHoursMinutes), _
                Color.BotInfo, " - ", Color.Message, numChars, Color.BotInfo, "/", Color.Message, maxPlayers, Color.BotInfo, " - Players: ", _
                Color.Message, charNames
        Case &H0: AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.Error, "No info received."
    End Select
End Sub

' MCP_CHARLOGON (0x07)
Private Sub rcvCharLogon()
Dim errMsg As String
    Select Case mcpPB.GetDWORD ' (DWORD) Result
        Case &H0
            AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.BotInfo, "Logged on character."
            sPB.SendRPacket &H12 ' 0x00: Success
        Case &H46: errMsg = "Player Not Found."
        Case &H7A: errMsg = "Logon Failed."
        Case &H7B: errMsg = "Character Expired."
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.Error, errMsg
End Sub

' MCP_CHARDELETE (0x0A)
Private Sub rcvCharDelete()
    Select Case mcpPB.GetDWORD ' (DWORD) Result
        ' (Diablo II v1.10 or later)
        Case &H0: AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.BotInfo, "Character deleted!"
        Case &H49: AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.Error, "Character doesn't exist."
    End Select
End Sub

' MCP_REQUESTLADDERDATA (0x11)
Private Sub rcvReqLadData()
Dim ladType As Long, totSize As Long, curSize As Long, totSizeUn As Long, rank As Long
Dim numEnt As Long, charExp() As Long, charFlags() As Long, charTitle() As Long, charLvl() As Long, charName() As Long
Dim i As Long
    With mcpPB
        ladType = .GetByte          '(BYTE) Ladder type
        'The 10-byte header:
        totSize = .GetWORD          '(WORD) Total response size
        curSize = .GetWORD          '(WORD) Current message size
        totSizeUn = .GetWORD        '(WORD) Total size of unreceived messages
        rank = .GetWORD             '(WORD) Rank of first entry
        .Skip 2                     '(WORD) Unknown (0)
        'Message Data:
        numEnt = .GetDWORD          '(DWORD)  Number of entries
        .Skip 4                     '(DWORD)  Unknown (0x10)
        'For each entry:
        For i = 0 To numEnt
            charExp(i) = .GetDWORD & .GetDWORD      '(QWORD) Character experience
            charFlags(i) = .GetByte                 '(BYTE) Character Flags
            charTitle(i) = .GetByte                 '(BYTE) Character title
            charLvl(i) = .GetWORD                   '(WORD) Character level
            charName(i) = .GetByte                  '(BYTE[16]) Character name
        Next i
    End With
    
    If BNET.varDebugMode = 1 Then
        AddChat Color.Carrot, "[Debug] ", Color.Bot, spID, Color.BotInfo, "Total Size: ", totSize, Color.Message, _
            Color.BotInfo, ", Current Size: ", Color.Message, curSize, Color.BotInfo, ", Total Unreceived: ", Color.Message, _
            totSizeUn, Color.BotInfo, ", Rank: ", Color.Message, rank, Color.BotInfo, ", Entries: ", Color.Message, numEnt
    End If

    For i = 0 To numEnt
        AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.BotInfo, "Name: ", Color.Message, charName(i), _
            Color.BotInfo, ", Exp: ", Color.Message, charExp(i), Color.BotInfo, ", Flags: ", Color.Message, charExp(i), _
            Color.BotInfo, ", Title: ", Color.Message, charTitle(i), Color.BotInfo, ", Level: ", Color.Message, charLvl
    Next i
End Sub

' MCP_MOTD (0x12)
Private Sub rcvMOTD()
    mcpPB.Skip 1 ' (BYTE) Unknown
    ' (STRING) MOTD
    AddChat Color.BotInfo, "[MCP] Realm MOTD: ", Color.Bot, spID, Color.BotInfo, mcpPB.GetString
End Sub

' MCP_CREATEQUEUE (0x14)
Private Sub rcvCreateQ()
    ' (DWORD) Position
    AddChat Color.BotInfo, "[MCP] Position in Queue: ", Color.Bot, spID, Color.BotInfo, mcpPB.GetDWORD
End Sub

' MCP_CHARLIST (0x17)
Private Sub rcvCharList()
Dim numCharReq As Long, numCharExist As Long, numCharRet As Long
Dim charName As String, charFlags As String, charStat As String
'Dim Count As String, pA() As String, pos As Long, tmpData As String
Dim i As Long
    With mcpPB
        numCharReq = .GetWORD       ' (WORD) Number of characters requested
        numCharExist = .GetDWORD    ' (DWORD) Number of characters that exist on this account
        numCharRet = .GetWORD       ' (WORD) Number of characters returned
        
        If numCharExist < 1 Then Exit Sub
        frmRealm.lstCharacter.ListItems.Clear
        For i = 1 To numCharRet
            ' For each character:
            charName = .GetString   ' (STRING) Name
            charFlags = .GetRaw(2)  ' (WORD) Flags
            charStat = .GetString   ' (STRING)Character statstring
            
            frmRealm.lstCharacter.ListItems.Add , , charName
            Call RealmCat(BNET.Product & "," & charName & "," & charFlags & charStat)
        Next i
    End With
    
    If BNET.varCRealm = 1 Then LogonRealmChar BNET.Character
End Sub

' MCP_CHARUPGRADE (0x18)
Private Sub rcvCharUpgrade()
Dim errMsg As String
    Select Case mcpPB
        Case &H0: AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.BotInfo, "Successfully converted character!"
        Case &H46: errMsg = "Character Not Found.": GoTo Failed
        Case &H7A: errMsg = "Upgrade failed.": GoTo Failed
        Case &H7B: errMsg = "Character is expired.": GoTo Failed
        Case &H7C: errMsg = "Already an expansion character.": GoTo Failed
    End Select
    Exit Sub
Failed:
    AddChat Color.BotInfo, svrMCP, Color.Bot, spID, Color.Error, errMsg
End Sub

' MCP_CHARLIST2 (0x19)
Private Sub rcvCharList2()
Dim numCharReq As Long, numCharExist As Long, numCharRet As Long
Dim charTime As Long, charName As String, charFlags As String, charStat As String
Dim i As Long
    With mcpPB
        numCharReq = .GetWORD       ' (WORD) Number of characters requested
        numCharExist = .GetDWORD    ' (DWORD) Number of characters that exist on this account
        numCharRet = .GetWORD       ' (WORD) Number of characters returned
        
        If numCharExist < 1 Then Exit Sub
        frmRealm.lstCharacter.ListItems.Clear
        For i = 1 To numCharRet
            ' For each character:
            charTime = .GetDWORD    ' (DWORD) Seconds since January 1 00:00:00 UTC 1970
            charName = .GetString   ' (STRING) Name
            charFlags = .GetRaw(2)  ' (WORD) Flags
            charStat = .GetString   ' (STRING)Character statstring
            
            frmRealm.lstCharacter.ListItems.Add , , charName
            Call RealmCat(BNET.Product & "," & charName & "," & charFlags & charStat)
        Next i
    End With
    
    If BNET.varCRealm = 1 Then LogonRealmChar BNET.Character
End Sub

' MCP_CHARCREATE C->S (0x02)
Public Function sndCharCr(c As String, xp As String, hc As String, n As String)
    d2charClass = charClass(c)
    With sPB
        .InsertDWORD "&H" & d2charClass         ' (DWORD) Character class
        .InsertWORD "&H" & charMask(xp, hc)     ' (WORD) Character flags
        .InsertNTString n                       ' (STRING) Character name
        .SendRPacket &H2                        ' MCP_CHARCREATE
    End With
End Function

' MCP_CREATEGAME (0x03)
Public Function sndCrGame(diff As String, Pdiff As String, maxP As String, gName As String, gPass As String, gDesc As String)
    d2gName = gName
    d2gPass = gPass
    ReqID = ReqID + 2
    With sPB
        .InsertWORD Val("&H" & ReqID)   ' (WORD) Request Id *
        .InsertDWORD CLng(diff)         ' (DWORD) Difficulty
        .InsertByte &H0                 ' (BYTE) Unknown - 1
        .InsertByte Val("&H" & Pdiff)   ' (BYTE) Player difference **
        .InsertByte Val("&H" & maxP)    ' (BYTE) Maximum players
        .InsertNTString gName           ' (STRING) Game name
        .InsertNTString gPass           ' (STRING) Game password
        .InsertNTString gDesc           ' (STRING) Game description
        .SendRPacket &H3
    End With
End Function

Public Function LogonRealmChar(charName As String)
    d2char = charName
    sPB.InsertNTString d2char
    sPB.SendRPacket &H7
End Function

Public Function DeleteChar(Character As String)
    With sPB
        .InsertWORD &H0
        .InsertNTString Character
        .SendRPacket &HA
        AddChat Color.BotInfo, "Character " & Character & " deleted."
    End With
End Function

Private Function charClass(Class As String) As String
    Select Case Class
        Case "Amazon": charClass = "00"
        Case "Sorceress": charClass = "01"
        Case "Necromancer": charClass = "02"
        Case "Paladin": charClass = "03"
        Case "Barbarian": charClass = "04"
        Case "Druid": charClass = "05"
        Case "Assassin": charClass = "06"
        Case Else: charClass = "00"
    End Select
End Function

Private Function charMask(xp As String, hc As String) As String
    If (xp = "1") And (hc = "1") Then
        charMask = "24"
    ElseIf (xp = "1") And (hc = "0") Then
        charMask = "20"
    ElseIf (xp = "0") And (hc = "1") Then
        charMask = "04"
    ElseIf (xp = "0") And (hc = "0") Then
        charMask = "00"
    Else
        charMask = "24"
    End If
End Function

Public Function PNMCP(p As Byte) As String
Select Case p
    Case &H1: PNMCP = "MCP_STARTUP"
    Case &H2: PNMCP = "MCP_CHARCREATE"
    Case &H3: PNMCP = "MCP_CREATEGAME"
    Case &H4: PNMCP = "MCP_JOINGAME"
    Case &H5: PNMCP = "MCP_GAMELIST"
    Case &H6: PNMCP = "MCP_GAMEINFO"
    Case &H7: PNMCP = "MCP_CHARLOGON"
    Case &HA: PNMCP = "MCP_CHARDELETE"
    Case &H11: PNMCP = "MCP_REQUESTLADDERDATA"
    Case &H12: PNMCP = "MCP_MOTD"
    Case &H13: PNMCP = "MCP_CANCELGAMECREATE"
    Case &H14: PNMCP = "MCP_CREATEQUEUE"
    Case &H16: PNMCP = "MCP_CHARRANK"
    Case &H17: PNMCP = "MCP_CHARLIST"
    Case &H18: PNMCP = "MCP_CHARUPGRADE"
    Case &H19: PNMCP = "MCP_CHARLIST2"
    Case Else: PNMCP = "UNKNOWN"
End Select
End Function
