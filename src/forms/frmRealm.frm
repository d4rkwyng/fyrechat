VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MsWinSck.ocx"
Begin VB.Form frmRealm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Realm Manager"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRealm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   5055
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock wsMCP 
      Left            =   3360
      Top             =   4320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Default         =   -1  'True
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
      Left            =   3960
      TabIndex        =   12
      Top             =   4320
      Width           =   945
   End
   Begin VB.CommandButton cmdDisconnect 
      BackColor       =   &H00808080&
      Caption         =   "Disconnect"
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
      Left            =   1200
      TabIndex        =   11
      ToolTipText     =   "Create Character"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdConnect 
      BackColor       =   &H00808080&
      Caption         =   "Connect"
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
      Left            =   120
      TabIndex        =   10
      ToolTipText     =   "Create Character"
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame fmChar 
      Caption         =   "Create Character"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   4815
      Begin VB.CommandButton cmdCreateChar 
         BackColor       =   &H00808080&
         Caption         =   "C&reate"
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
         Left            =   3600
         TabIndex        =   9
         ToolTipText     =   "Create Character"
         Top             =   840
         Width           =   975
      End
      Begin VB.CheckBox chkExpChar 
         Caption         =   "Expansion"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox chkHardChar 
         Caption         =   "Hardcore"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1215
      End
      Begin VB.ComboBox lstClass 
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
         ItemData        =   "frmRealm.frx":000C
         Left            =   2760
         List            =   "frmRealm.frx":0025
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox txtCreateChar 
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
         Left            =   240
         MaxLength       =   15
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Realm 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton cmdLogon 
         BackColor       =   &H00808080&
         Caption         =   "&Logon"
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
         Left            =   3600
         TabIndex        =   4
         ToolTipText     =   "Logon to Character"
         Top             =   2040
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
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
         Left            =   2520
         TabIndex        =   3
         ToolTipText     =   "Delete Character"
         Top             =   2040
         Width           =   975
      End
      Begin MSComctlLib.ListView lstCharacter 
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   2990
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Character"
            Object.Width           =   2716
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   2364
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Lvl"
            Object.Width           =   811
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "HC"
            Object.Width           =   741
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Exp"
            Object.Width           =   882
         EndProperty
      End
   End
End
Attribute VB_Name = "frmRealm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.icon = frmMain.icon
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer): Call ExitProgram(Cancel): End Sub
Private Sub cmdClose_Click(): Me.Hide: End Sub
Private Sub lstCharacter_DblClick(): LogonRealmChar lstCharacter.SelectedItem: End Sub

Private Sub cmdDisconnect_Click()
    RealmDisconnect
End Sub

Public Sub RealmDisconnect()
    AddChat Color.Error, "Disconnected from MCP."
    Call CloseRealm
End Sub
Private Sub cmdConnect_Click()
    If cmdConnect.Enabled Then
        sndLogRealmEx BNET.Realm
    End If
End Sub

Private Sub cmdCreateChar_Click()
    If wsMCP.State = 0 Then
        MsgBox "Currently not connected to a realm."
    Else
        Select Case lstClass.text
            Case "Assassin", "Druid"
                If Not BNET.Product = "PX2D" Then
                    MsgBox "Cannot create an expansion character with regular Diablo 2."
                    Exit Sub
                ElseIf chkExpChar = 0 Then
                    chkExpChar.Value = 1
                End If
        End Select
        sndCharCr lstClass.text, chkExpChar.Value, chkHardChar.Value, txtCreateChar.text
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim Msg, DChar As String
    Dim Cancel As Boolean
    If wsMCP.State = 0 Then
        MsgBox "Currently not connected to a realm."
    Else
        DChar = lstCharacter.SelectedItem
        If lstCharacter.SelectedItem.index > 0 Then
            Msg = "Are you sure you want to delete?"
        If MsgBox(Msg, vbQuestion + vbYesNo, "Delete " & DChar) = vbNo Then
            Cancel = True
            Exit Sub
        End If
        lstCharacter.ListItems.Remove lstCharacter.SelectedItem.index
        DeleteChar DChar
        End If
    End If
End Sub

Private Sub cmdLogon_Click()
    If wsMCP.State = 0 Then
        MsgBox "Currently not connected to a realm."
    Else
        Dim SelectItem As String
        SelectItem = lstCharacter.SelectedItem
        If Len(SelectItem) > 3 Then
            LogonRealmChar lstCharacter.SelectedItem
        Else
            MsgBox "No character selected"
        End If
    End If
End Sub

Public Sub CloseRealm()
    lstCharacter.ListItems.Clear
    wsMCP.Close
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
End Sub

Private Sub ExitProgram(Cancel As Integer)
    If Urealm Then
        frmRealm.Hide
        Cancel = True
    End If
End Sub


Private Sub wsMCP_Connect()
On Local Error Resume Next
    wsMCP.SendData Chr(1)
    AddChat Color.BotInfo, "[MCP] Connected to MCP!"
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    With sPB
        .InsertDWORD rCookie
        .InsertDWORD rStatus
        .InsertNonNTString rChunk1
        .InsertNonNTString rChunk2
        .InsertNTString RealmName
        .SendRPacket &H1
    End With
End Sub

Private Sub wsMCP_DataArrival(ByVal bytesTotal As Long)
Dim TempData As String
    wsMCP.GetData TempData, vbString
    ParseMCP TempData
End Sub

Private Sub wsMCP_Close()
    AddChat Color.BotInfo, svrMCP, Color.BotError, "Connection to MCP has been lost!"
    Call CloseRealm
End Sub

Private Sub wsMCP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    AddChat Color.BotInfo, svrMCP, Color.Error, "Error [" & Number & "] " & Description
    wsMCP.Close
End Sub
