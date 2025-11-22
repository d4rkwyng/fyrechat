Attribute VB_Name = "modWebList"
Option Explicit

Public Sub WebListAdd(ByVal username As String, ByVal Product As String, ByVal Flags As Long, ByVal Ping As Long)
On Local Error Resume Next
    Dim icon As String
    Dim nPing
    Dim i As Integer
    Dim Found As Boolean
    Dim ProdID As String
    icon = ""
    
    ProdID = StrReverse(Left$(Product, 4))
    nPing = Ping
    
    If (Flags And BNFLAGS_BLIZZ) Then ProdID = "BLIZZ"
    If (Flags And BNFLAGS_SYSOP) Then ProdID = "SYSOP"
    If (Flags And BNFLAGS_OP) Then ProdID = "OP"
    If (Flags And BNFLAGS_SPKR) Then ProdID = "SPKR"
    If (Flags And BNFLAGS_GLASSES) Then ProdID = "GLASSES"
    If (Flags And BNFLAGS_SQUELCH) Then ProdID = "SQUELCH"
    If (Flags And BNFLAGS_GFOFFICIAL) Then ProdID = "GFOFFICIAL"
    If (Flags And BNFLAGS_GFPLAYER) Then ProdID = "GFPLAYER"
    If (Flags And BNFLAGS_PLUG) Then nPing = "PLUG"
    
    Dim tmpUser() As String
    Dim tmpUser2() As String
    For i = 1 To frmMain.lstWebChan.ListItems.Count
    '[ICON]SEXP[/ICON][USER]venox[xL][/USER][PING]-1[/PING]
        tmpUser() = Split(frmMain.lstWebChan.ListItems(i).text, "[USER]")
        tmpUser2() = Split(tmpUser(1), "[/USER]")
        If LCase(tmpUser2(0)) = LCase(username) Then
        'If InStr(LCase(frmMain.lstWebChan.ListItems(i).text), LCase(UserName)) Then
            Found = True
            Exit For
        Else
            Found = False
        End If
    Next i
    If Found Then frmMain.lstWebChan.ListItems.Remove (i)
    
    Dim UserData As String
    UserData = "[ICON]" & ProdID & "[/ICON][USER]" & username & "[/USER][PING]" & nPing & "[/PING]"
    If (ProdID = "OP") Then
        frmMain.lstWebChan.ListItems.Add 1, , UserData
    Else
        frmMain.lstWebChan.ListItems.Add , , UserData
    End If
End Sub

Public Sub WebListDel(ByVal username As String)
    Dim i As Integer
    Dim Found As Boolean
    Dim tmpUser() As String
    Dim tmpUser2() As String
    For i = 1 To frmMain.lstWebChan.ListItems.Count
        tmpUser() = Split(frmMain.lstWebChan.ListItems(i).text, "[USER]")
        tmpUser2() = Split(tmpUser(1), "[/USER]")
        If (LCase(tmpUser2(0)) = LCase(username)) Then
            Found = True
            Exit For
        Else
            Found = False
        End If
    Next i
    If Found Then
        frmMain.lstWebChan.ListItems.Remove (i)
    End If
End Sub

Public Sub WriteWebList()
On Local Error Resume Next
    Dim i As Integer
    ' Delete File
    Kill (App.Path & "/channel_list.txt")
    
    Dim fHandle As Long
    fHandle = FreeFile()
    Call IsPathAFolder(App.Path)
    If PathFolder Then
        Open (App.Path & "/channel_list.txt") For Append As #fHandle
    End If
    For i = 1 To frmMain.lstWebChan.ListItems.Count
        Print #fHandle, frmMain.lstWebChan.ListItems(i).text
    Next i
    Close #fHandle
End Sub

Public Sub WebListUpload()
    Dim UploadResult As Boolean
    UploadResult = UploadFile(strWL_FTPAddress, strWL_FTPUser, strWL_FTPPass, _
        App.Path & "\channel_list.txt", strWL_FTPLoc & "channel_list.txt")
    If UploadResult Then
        If UpBool Then: MsgBox "Success!"
    Else
        AddChat Color.Error, "Failed to upload channel_list.txt"
        If UpBool Then MsgBox "Failed!"
    End If
End Sub

Private Function UploadFile(ByVal HostName As String, _
    ByVal username As String, _
    ByVal password As String, _
    ByVal LocalFileName As String, _
    ByVal RemoteFileName As String) As Boolean

    Dim FTP As Inet
    Set FTP = New Inet
    With FTP
        .Protocol = icFTP
        .RemoteHost = HostName
        .username = username
        .password = password
        .Execute .URL, "Put " + LocalFileName + " " + RemoteFileName
        Do While .StillExecuting
            DoEvents
        Loop
        UploadFile = (.ResponseCode)
    End With
    Set FTP = Nothing
End Function

