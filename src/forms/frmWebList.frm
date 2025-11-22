VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmWebList 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Web List"
   ClientHeight    =   4575
   ClientLeft      =   1980
   ClientTop       =   6015
   ClientWidth     =   6420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkWLEnabled 
      Caption         =   "On/Off"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   14
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   13
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdUpload 
      Caption         =   "Upload"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox txtUpTimer 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2640
      MaxLength       =   15
      TabIndex        =   9
      Text            =   "60"
      Top             =   3480
      Width           =   615
   End
   Begin VB.TextBox txtFTPLoc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      MaxLength       =   45
      TabIndex        =   6
      Text            =   "public_html/"
      Top             =   3120
      Width           =   1935
   End
   Begin VB.TextBox txtFTPPass 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      MaxLength       =   20
      TabIndex        =   8
      Top             =   3120
      Width           =   1815
   End
   Begin VB.TextBox txtFTPUser 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4440
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtFTPAddress 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1320
      MaxLength       =   60
      TabIndex        =   5
      Text            =   "domain.com"
      Top             =   2760
      Width           =   1935
   End
   Begin MSComctlLib.ListView lstWebChan 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      OLEDragMode     =   1
      _Version        =   393217
      Icons           =   "ImgLst"
      SmallIcons      =   "ImgLst"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ChanSave"
         Object.Width           =   10583
      EndProperty
   End
   Begin VB.Label lblUTimer 
      Caption         =   "Upload Timer (seconds)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   3480
      Width           =   2295
   End
   Begin VB.Label lblFTPLoc 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblFTPPass 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   3480
      TabIndex        =   4
      Top             =   3120
      Width           =   855
   End
   Begin VB.Label lblFTPUser 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3480
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label lblFTPDomain 
      Caption         =   "FTP Domain"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   2760
      Width           =   1215
   End
End
Attribute VB_Name = "frmWebList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, _
ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" _
(ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, _
ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, _
ByVal lContext As Long) As Long

Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, _
ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer

Private Const INTERNET_SERVICE_FTP = 1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2

Public Function UploadFile(ByRef server As String, ByRef username As String, ByRef password As String, _
ByRef port As Integer, ByRef localFileName As String, ByRef remoteFileName As String) As Boolean
    Dim session As Long
    Dim connection As Long
    
    On Error GoTo UploadError
    
    session = InternetOpen("FyreChat", 0, vbNullString, vbNullString, 0)
    connection = InternetConnect(session, server, port, username, password, INTERNET_SERVICE_FTP, 0, 0)
    
    If FtpPutFile(connection, localFileName, remoteFileName, FTP_TRANSFER_TYPE_BINARY, 0) Then
        UploadFile = True
    Else
        UploadFile = False
    End If
    
    InternetCloseHandle session
    InternetCloseHandle connection
    
    Exit Function
UploadError:
    AddChat Color.BotInfo, "[Bot] ", Color.Error, "Upload Function Error: " & Err.Description
 End Function
 
Private Sub cmdUpload_Click()
    WebListSave
    If frmWebList.UploadFile(strWL_FTPAddress, strWL_FTPUser, strWL_FTPPass, 21, _
        App.Path & "\channel_list.txt", strWL_FTPLoc & "channel_list.txt") Then
        MsgBox "Upload Completed"
    Else
        MsgBox "Error Uploading"
    End If
End Sub

Private Sub Form_Load()
    Me.icon = frmMain.icon
    chkWLEnabled.Value = blWL_Enabled
    txtUpTimer.text = intWL_UpTimer
    txtFTPAddress.text = strWL_FTPAddress
    txtFTPUser.text = strWL_FTPUser
    txtFTPPass.text = strWL_FTPPass
    txtFTPLoc.text = strWL_FTPLoc
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmMain.wsBNET.State = 7 Then
        Cancel = True
        Me.Hide
    End If
End Sub


Private Sub cmdSave_Click()
    Dim tmpWeb As Long, tmpUpload As Long
    blWL_Enabled = chkWLEnabled
    
    If Not txtUpTimer.text >= 10 Then
        AddChat Color.BotInfo, "[Bot] ", Color.Error, "Unable to save FTP Upload time; Replacing with lowest interval (10)."
        txtUpTimer.text = 10
    End If
    intWL_UpTimer = txtUpTimer.text
    
    If Not Len(txtFTPAddress.text) >= 3 Then
        AddChat Color.BotInfo, "[Bot] ", Color.Error, "Unable to save domain (size too small)."
    Else
        strWL_FTPAddress = txtFTPAddress.text
    End If
    
    If Not Len(txtFTPUser.text) >= 1 Then
        AddChat Color.BotInfo, "[Bot] ", Color.Error, "Unable to save user (size too small)."
    Else
        strWL_FTPUser = txtFTPUser.text
    End If
    
    If Not Len(txtFTPPass.text) >= 1 Then
        AddChat Color.BotInfo, "[Bot] ", Color.Error, "Unable to save password (size too small)."
    Else
        strWL_FTPPass = txtFTPPass.text
    End If
    
    If Not Mid(txtFTPLoc.text, Len(txtFTPLoc.text), 1) = "/" Then
        AddChat Color.BotInfo, "[Bot] ", Color.Bot, "Automatically adding closing slash to FTP Location."
        strWL_FTPLoc = txtFTPLoc.text & "/"
        txtFTPLoc.text = strWL_FTPLoc
    Else
        strWL_FTPLoc = txtFTPLoc.text
    End If
    
    modFunctions.SaveConfig
End Sub

Private Sub cmdClose_Click()
    frmWebList.Hide
End Sub

Public Sub WebListAdd(ByVal username As String, ByVal Product As String, ByVal flags As Long, ByVal ping As Long)
Dim i As Integer, STAT As String, nPing As String
Dim Found As Boolean, strWebList As String
On Local Error Resume Next
    
    STAT = StrReverse(Left$(Product, 4))
    nPing = CStr(ping)
    
    If (flags And BNFLAGS_BLIZZ) Then STAT = "BLIZZ"
    If (flags And BNFLAGS_SYSOP) Then STAT = "SYSOP"
    If (flags And BNFLAGS_OP) Then STAT = "OP"
    If (flags And BNFLAGS_SPKR) Then STAT = "SPKR"
    If (flags And BNFLAGS_GLASSES) Then STAT = "GLASSES"
    If (flags And BNFLAGS_SQUELCH) Then STAT = "SQUELCH"
    If (flags And BNFLAGS_GFOFFICIAL) Then STAT = "GFOFFICIAL"
    If (flags And BNFLAGS_GFPLAYER) Then STAT = "GFPLAYER"
    If (flags And BNFLAGS_PLUG) Then nPing = "PLUG"
    
    Dim tmpUser() As String
    Dim tmpUser2() As String
    For i = 1 To lstWebChan.ListItems.count
    '[ICON]SEXP[/ICON][USER]venox[xL][/USER][PING]-1[/PING]
        tmpUser() = Split(lstWebChan.ListItems(i).text, "[USER]")
        tmpUser2() = Split(tmpUser(1), "[/USER]")
        If LCase(tmpUser2(0)) = LCase(username) Then
            Found = True
            Exit For
        Else
            Found = False
        End If
    Next i
    If Found Then lstWebChan.ListItems.Remove (i)
    
    strWebList = "[ICON]" & STAT & "[/ICON][USER]" & username & "[/USER][PING]" & nPing & "[/PING]"
    
    If (STAT = "OP") Then
        lstWebChan.ListItems.Add 1, , strWebList
    Else
        lstWebChan.ListItems.Add , , strWebList
    End If
End Sub

Public Sub WebListDel(ByVal username As String)
On Local Error Resume Next
    Dim i As Integer, Found As Boolean
    Dim tmpUser() As String, tmpUser2() As String
    
    For i = 1 To lstWebChan.ListItems.count
        tmpUser() = Split(lstWebChan.ListItems(i).text, "[USER]")
        tmpUser2() = Split(tmpUser(1), "[/USER]")
        If (LCase(tmpUser2(0)) = LCase(username)) Then
            Found = True
            Exit For
        Else
            Found = False
        End If
    Next i
    
    If Found Then lstWebChan.ListItems.Remove (i)

    Erase tmpUser(): Erase tmpUser2()
End Sub

Public Sub RefreshWebChanList(username As String, Product As String, flags As Long, ping As Long)
Dim i As Integer, STAT As String, nPing As String
Dim tmpUser() As String, tmpUser2() As String
Dim tmpPing() As String, tmpPing2() As String
Dim strWebList As String
    On Local Error Resume Next
    STAT = StrReverse(Left$(Product, 4))
    nPing = CStr(ping)
    
    If (flags And BNFLAGS_BLIZZ) Then STAT = "BLIZZ"
    If (flags And BNFLAGS_SYSOP) Then STAT = "SYSOP"
    If (flags And BNFLAGS_OP) Then STAT = "OP"
    If (flags And BNFLAGS_SPKR) Then STAT = "SPKR"
    If (flags And BNFLAGS_GLASSES) Then STAT = "GLASSES"
    If (flags And BNFLAGS_SQUELCH) Then STAT = "SQUELCH"
    If (flags And BNFLAGS_GFOFFICIAL) Then STAT = "GFOFFICIAL"
    If (flags And BNFLAGS_GFPLAYER) Then STAT = "GFPLAYER"
    If (flags And BNFLAGS_PLUG) Then nPing = "PLUG"
    
    For i = 1 To lstWebChan.ListItems.count
        tmpUser() = Split(lstWebChan.ListItems(i).text, "[USER]")
        tmpUser2() = Split(tmpUser(1), "[/USER]")
        
        If username = tmpUser2(0) Then
            lstWebChan.ListItems.Remove (i)
            tmpPing() = Split(lstWebChan.ListItems(i).text, "[PING]")
            tmpPing2() = Split(tmpPing(1), "[/PING]")
            nPing = tmpPing2(0)
            strWebList = "[ICON]" & STAT & "[/ICON][USER]" & username & "[/USER][PING]" & nPing & "[/PING]"
            If (STAT = "OP") Then
                lstWebChan.ListItems.Add 1, , strWebList
            Else
                lstWebChan.ListItems.Add , , strWebList
            End If
        End If
    Next i

    Erase tmpUser(): Erase tmpUser2()
    Erase tmpPing(): Erase tmpPing2()
End Sub

Public Sub WebListSave()
On Error GoTo SaveError
    Dim i As Integer, fHand As Long
    
    If FileExists(App.Path & "\channel_list.txt") Then Kill App.Path & "\channel_list.txt"
    fHand = FreeFile()
    Open App.Path & "\channel_list.txt" For Output As #fHand
        For i = 1 To lstWebChan.ListItems.count
            Print #fHand, lstWebChan.ListItems(i).text
        Next i
    Close #fHand
    fHand = 0
    Exit Sub
SaveError:
    AddChat Color.Error, "Save Error #: " & Err.Number & ", " & _
        "Error Desc: " & Err.Description
    Close #fHand
    Exit Sub
End Sub

