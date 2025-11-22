Attribute VB_Name = "modIcons"
Option Explicit
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private ICONFile                As String
Private PINGIconFile(1 To 9)    As String
Private BNIFILEEXISTS           As Boolean
Private ICOFILEEXISTS            As Boolean
Private BnetIcon()              As xBnetIcon
Private BnetPing(1 To 9)        As xBnetPing
Private ChannelString           As String
Private NumberOfIcons           As Long

Private Type xBnetIcon
    IconID  As Long
    x       As Long
    y       As Long
    Tag     As String
    unknown As Long
    flags   As Long
    IListID As Integer
    ThePic  As String
End Type

Public Type xBnetPing
    IListID As Integer
    High    As Long
    Low     As Long
    flag    As Long
End Type

Private mclsToolTip As New clsToolTip

Public Sub SetupChannel(sdata As String)
    ChannelString = sdata
    frmMain.lstChannel.ListItems.Clear
End Sub

Public Function AddUser(user As String, Product As String, flags As Long, ping As Long)
On Local Error Resume Next
    Dim fcolor As String, ParsedString As String
    Call ParseStatString(Product, ParsedString)
    Product = Left$(Product, 4)
    With frmMain.lstChannel
        If (flags And BNFLAGS_OP) Then
            .ListItems.Add 1, , user, , GetBnetIconCode(Product, flags)
            .ListItems(1).ListSubItems.Add 1, , , GetBnetPingCode(ping, flags), ping
            If ParsedString = "" Then
                .ListItems(1).ToolTipText = "(Operator) " & tmpFlag
            Else
                .ListItems(1).Tag = ParsedString
                .ListItems(1).ToolTipText = "(Operator) " & ParsedString
            End If
            fcolor = Color.Op
            .ListItems(1).ForeColor = fcolor
        Else
            .ListItems.Add , , user, , GetBnetIconCode(Product, flags)
            .ListItems(.ListItems.count).ListSubItems.Add , , , GetBnetPingCode(ping, flags), ping
            If ParsedString = "" Then
            Else
                .ListItems(.ListItems.count).Tag = ParsedString
                .ListItems(.ListItems.count).ToolTipText = ParsedString
            End If
            fcolor = Color.user
            If user = BNET.TrueUsername Then fcolor = Color.Self
            .ListItems(.ListItems.count).ForeColor = fcolor
        End If
        If user = BNET.TrueUsername Then
            MyFlags = flags
            MyPing = ping
            MyParsedString = ParsedString
        End If
    End With
    
    
    'mclsToolTip.ToolText(.ListItems(1)) = "Product: " & ParsedString & vbCrLf & _
    '    "Ping: " & Ping & vbCrLf & _
    '    "Flags: " & Flags
    'mclsToolTip.AddTool .ListItems(1)
    'With mclsToolTip
    '    .RemoveTool frmMain.txtChanName
    '    .ToolTipHeaderShow = False
    '    .ToolText(frmMain.txtChanName) = "Logged on to battle.net as " & _
    '    Bnet.TrueUsername & ":" & vbCrLf & _
    '    "Product: " & MyParsedString & vbCrLf & _
    '    "Ping: " & MyPing & vbCrLf & _
    '    "Flags: " & MyFlags & vbCrLf & _
    '    "Channel: " & Bnet.CurrentChan & vbCrLf & _
    '    "Users: " & frmMain.lstChannel.ListItems.Count
    '    .AddTool frmMain.txtChanName
    'End With
    frmMain.txtChanName.text = ChannelString & " (" & frmMain.lstChannel.ListItems.count & ")"
End Function

Public Function RemoveUser(user As String)
On Local Error Resume Next
    With frmMain
        .txtChanName.Clear
        .txtChanName.AddItem ChannelString & " (" & .lstChannel.ListItems.count - 1 & ")"
        .lstChannel.ListItems.Remove .lstChannel.FindItem(user).index
        .txtChanName.text = ChannelString & " (" & .lstChannel.ListItems.count & ")"
    End With
End Function

Public Sub RefreshChannelList(user As String, Product As String, flags As Long)
On Local Error Resume Next
    Dim i As Integer
    With frmMain
        For i = 1 To .lstChannel.ListItems.count
            If .lstChannel.ListItems.item(i).text = user Then .lstChannel.ListItems.item(i).SmallIcon = GetBnetIconCode(Product, flags)
        Next i
    End With
End Sub

Public Function SetIconPaths()
On Local Error Resume Next
    ICONFile = App.Path & "\icons.bni"
    PINGIconFile(1) = App.Path & "\ping\1b.bmp"
    PINGIconFile(2) = App.Path & "\ping\0.bmp"
    PINGIconFile(3) = App.Path & "\ping\1g.bmp"
    PINGIconFile(4) = App.Path & "\ping\2g.bmp"
    PINGIconFile(5) = App.Path & "\ping\3y.bmp"
    PINGIconFile(6) = App.Path & "\ping\4y.bmp"
    PINGIconFile(7) = App.Path & "\ping\5r.bmp"
    PINGIconFile(8) = App.Path & "\ping\6r.bmp"
    PINGIconFile(9) = App.Path & "\ping\plug.bmp"
End Function

Public Function InitPingIcons()
On Local Error Resume Next
    Dim i As Long
    Dim i2 As Long
    
    BnetPing(9).flag = 22
    
    If NumberOfIcons = 0 Then NumberOfIcons = 1
    
    For i = 1 To UBound(PINGIconFile)
        frmMain.addPic.Picture = LoadPicture(PINGIconFile(i))
        frmMain.BnetIcons.ListImages.Add NumberOfIcons, , frmMain.addPic.Image
        BnetPing(i).IListID = NumberOfIcons
    
        NumberOfIcons = NumberOfIcons + 1
        Set frmMain.addPic = Nothing
        DoEvents
        Dim strta As String
        strta = BnetPing(i).flag & " " & BnetPing(i).High & " " & BnetPing(i).Low
        frmMain.lstChannel.ListItems.Add , , strta, , BnetPing(i).IListID 'was ,1 , 1
    Next i

    BnetPing(1).Low = -1: BnetPing(1).High = 0 'negative 1
    BnetPing(2).Low = 0: BnetPing(2).High = 4 '0ms
    BnetPing(3).Low = 5: BnetPing(3).High = 199 '1g
    BnetPing(4).Low = 200: BnetPing(4).High = 300 '2g
    BnetPing(5).Low = 301: BnetPing(5).High = 400 '3y
    BnetPing(6).Low = 401: BnetPing(6).High = 500 '4y
    BnetPing(7).Low = 501: BnetPing(7).High = 600 '5r
    BnetPing(8).Low = 601: BnetPing(8).High = 1200 '6r
    BnetPing(9).Low = &H0: BnetPing(9).High = &H0 'PLUG status
End Function

Public Function CheckBnetIcons()
On Local Error Resume Next
    Dim iconscheck As String
    iconscheck = Dir(ICONFile)
    
    If iconscheck = "" Then
        BNIFILEEXISTS = False
        Exit Function
    Else
        BNIFILEEXISTS = True
        Call InitIcons
    End If
End Function

Public Function CheckPingIcons()
On Local Error Resume Next
    Dim iconscheck As String
    Dim i
    For i = 1 To UBound(PINGIconFile)
        iconscheck = Dir(PINGIconFile(i))
        
        If iconscheck = "" Then
            ICOFILEEXISTS = False
            Exit Function
        Else
            ICOFILEEXISTS = True
        End If
    Next i
    Call InitPingIcons
End Function

Public Function InitIcons()
On Local Error Resume Next
    Dim strFileName As String, strBuffer As String
    Dim dwCount As Long, cPos As Long

    strFileName = ICONFile
    
    Open strFileName For Binary As #1
        strBuffer = Input(LOF(1), #1)
    Close #1
    
    dwCount = ("&H" & StrToHex(StrReverse(Mid(strBuffer, 9, 4))))
    cPos = 17 '17
    
    ReDim BnetIcon(1 To dwCount) As xBnetIcon
    Dim i
    For i = 1 To dwCount
        BnetIcon(i).IconID = ("&H" & StrToHex(StrReverse((Mid(strBuffer, cPos, 4)))))
        cPos = cPos + 4
        BnetIcon(i).x = ("&H" & StrToHex(StrReverse(Mid(strBuffer, cPos, 4))))
        cPos = cPos + 4
        BnetIcon(i).y = ("&H" & StrToHex(StrReverse(Mid(strBuffer, cPos, 4))))
        cPos = cPos + 4
        BnetIcon(i).Tag = Mid(strBuffer, cPos, 4)
        
        cPos = cPos + 4
        BnetIcon(i).unknown = ("&H" & StrToHex(StrReverse(Mid(strBuffer, cPos, 4))))
        'cPos = cPos + 4 'Used in icons_star.bni only
        
        If i >= 9 Then      'Do not use this if statement
            cPos = cPos + 4 'for icons_star must add 4 after
        End If              '.unknown regardless
        
        If InStr(BnetIcon(i).Tag, vbNullChar) Then: BnetIcon(i).Tag = "None"
        
        BnetIcon(i).IListID = i
        frmMain.BnetIcons.MaskColor = vbBlack
    Next i

    Dim tgaFile As New clsBNI
    Dim Scaling As Single
    Scaling = 1
    frmMain.Picture1.Cls
    tgaFile.Autoscale = False
    tgaFile.LoadTGA cPos, strFileName
    frmMain.Picture1.Width = tgaFile.TGAWidth
    frmMain.Picture1.Height = tgaFile.TGAHeight + 67.1
    Call tgaFile.DrawTGA(frmMain.Picture1)
    DoEvents

    Call SetImages
End Function

Private Function GetBnetIconCode(ByVal statstring As String, ByVal flags As Long) As Integer
On Local Error Resume Next
    Dim strstProduct As String
    Dim Values() As String
    Dim flId As Long
    Dim stproduct
    Dim i
    Values() = Split(Mid(statstring, 5, Len(statstring)))
    stproduct = Mid$(statstring, 1, 4)

    If BNIFILEEXISTS = True Then
        Select Case flags
            Case 22, 16:
            Case 0, 22:
            Case Else:
                For i = 1 To UBound(BnetIcon)
                    flId = BnetIcon(i).IconID
                    If (flId And flags) = flId Then
                        GetBnetIconCode = BnetIcon(i).IListID
                        Exit Function
                    End If
                Next i
        End Select
    
        For i = 1 To UBound(BnetIcon)
            If stproduct = BnetIcon(i).Tag Then
                GetBnetIconCode = BnetIcon(i).IListID
                Exit Function
            End If
        Next i
    End If
    GetBnetIconCode = 0
End Function

Private Function GetBnetPingCode(ByVal ping As Long, ByVal flags As Long) As Integer
On Local Error Resume Next
    Dim i
    If ICOFILEEXISTS = True Then
        If (flags And BNFLAGS_PLUG) = BNFLAGS_PLUG Then
            GetBnetPingCode = BnetPing(9).IListID
            Exit Function
        End If
        Select Case ping
            Case Is = -1: GetBnetPingCode = BnetPing(1).IListID
            Case Is = 0: GetBnetPingCode = BnetPing(2).IListID
            Case 0 To 199: GetBnetPingCode = BnetPing(3).IListID
            Case 200 To 299: GetBnetPingCode = BnetPing(4).IListID
            Case 300 To 399: GetBnetPingCode = BnetPing(5).IListID
            Case 400 To 499: GetBnetPingCode = BnetPing(6).IListID
            Case 500 To 599: GetBnetPingCode = BnetPing(7).IListID
            Case Is > 600: GetBnetPingCode = BnetPing(8).IListID
            Case Else: GetBnetPingCode = 0
        End Select
    End If
End Function

Private Function StrToHex(ByVal String1 As String) As String
    Dim strTemp As String, strReturn As String, i As Long
    For i = 1 To Len(String1)
        strTemp = Hex(Asc(Mid$(String1, i, 1)))
        If Len(strTemp) = 1 Then strTemp = "0" & strTemp
        strReturn = strReturn & strTemp
    Next i
    StrToHex = strReturn
End Function

Private Function SetImages()
On Local Error Resume Next
    Dim AddThis As Integer, x2 As Integer, x3 As Integer
    Dim pos As Integer, pos2 As Integer, pos3 As Integer, pos4 As Integer
    pos = 0
    pos2 = 0
    pos3 = 14
    pos4 = 32
    AddThis = 14
    Dim x
    Dim y
    Dim i
    For x = 1 To UBound(BnetIcon)
        For y = pos To pos3
            For i = pos2 To pos4
                SetPixel frmMain.addPic.hdc, x2, x3, GetPixel(frmMain.Picture1.hdc, i, y)
                x2 = x2 + 1
            Next i
            x2 = 0
            x3 = x3 + 1
        Next y
        pos = pos + AddThis
        pos3 = pos3 + AddThis
        x3 = 0
        x2 = x2
        frmMain.BnetIcons.ListImages.Add x, , frmMain.addPic.Image
        Set frmMain.addPic = Nothing
    Next x
    NumberOfIcons = x
End Function
