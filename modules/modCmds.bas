Attribute VB_Name = "modCmds"
Option Explicit

Public Sub ProcessBotCommand(ByVal cmd As String, ByVal usr As String)
On Local Error Resume Next
    Dim spltCmd() As String, cmdProc As String
    Dim i As Long
    
    If InStr(cmd, " ") Then
        spltCmd = Split(cmd, " ", 2)
        cmdProc = Mid(spltCmd(0), 2)
    Else
        cmdProc = Mid(cmd, 2)
    End If
    
    If frmMain.wsBNET.State = sckConnected Then
        Select Case LCase(cmdProc)
            Case "reply": Send "/w " & WReply & " " & spltCmd(1), frmMain.wsBNET
            Case "rejoin": Rejoin BNET.CurrentChan
            Case "clearscreen": frmMain.Chat_Output.text = ""
            Case "clear": ClearBuffers: Exit Sub
            Case "clearqueue": frmMain.lstEnqueue.Clear
            Case "disconnect": frmMain.DisconnectProc
            Case "reconnect": frmMain.ConnectProc
            Case "compuptime": Send App.ProductName & " - Computer Uptime: " & FormatCount(GetTickCount), frmMain.wsBNET
            Case "botuptime": Send App.ProductName & " Uptime: " & FormatCount(GetTickCount - RunningTime), frmMain.wsBNET
            Case "connected": Send App.ProductName & " - Connection Uptime: " & FormatCount(GetTickCount - connecttime), frmMain.wsBNET
            Case "downloadfile"
                blDownloadFile = True
                With sPB
                    .InsertDWORD &H1            ' (DWORD) Request ID
                    .InsertDWORD &H0            ' (DWORD) Unknown
                    .InsertNTString spltCmd(1)  ' (STRING) Filename
                    .SendPacket &H33            ' SID_GETFILETIME
                End With
            Case "listchannels"
                Dim chanLst As String
                chanLst = frmMain.lstChanSave.ListItems.item(1).text
                For i = 2 To frmMain.lstChanSave.ListItems.count
                    chanLst = chanLst & ", " & frmMain.lstChanSave.ListItems.item(i).text
                Next i
                AddChat Color.BotInfo, "[Bot] Channel List: ", Color.Message, chanLst
            Case Else: Send cmd, frmMain.wsBNET
        End Select
        Exit Sub
    End If
    
    Select Case cmdProc
        Case "downloadfile": ConnectBNETFTP spltCmd(1)
    End Select

    Erase spltCmd()
End Sub

