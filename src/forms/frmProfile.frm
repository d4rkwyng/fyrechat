VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmProfile 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "User's Profile"
   ClientHeight    =   5070
   ClientLeft      =   8100
   ClientTop       =   6210
   ClientWidth     =   4635
   FillColor       =   &H80000012&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmProfile.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   4635
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H00000000&
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
      Left            =   3360
      MaskColor       =   &H00000000&
      TabIndex        =   6
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00000000&
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
      Left            =   2160
      MaskColor       =   &H00000000&
      TabIndex        =   5
      Top             =   4440
      Width           =   1095
   End
   Begin RichTextLib.RichTextBox rtbDescription 
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   4471
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmProfile.frx":000C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbLocation 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1080
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmProfile.frx":0087
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbUsername 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmProfile.frx":0102
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbSex 
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmProfile.frx":017D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtbAge 
      Height          =   375
      Left            =   3480
      TabIndex        =   11
      Top             =   360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmProfile.frx":01F8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblDescription 
      Caption         =   "Description:"
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
      TabIndex        =   4
      ToolTipText     =   "Description"
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label lblLocation 
      Caption         =   "Location:"
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
      TabIndex        =   3
      ToolTipText     =   "Location"
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblAge 
      Caption         =   "Age:"
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
      Left            =   3480
      TabIndex        =   2
      ToolTipText     =   "Age"
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblSex 
      Caption         =   "Sex:"
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
      Left            =   2040
      TabIndex        =   1
      ToolTipText     =   "Sex"
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblName 
      Caption         =   "Name:"
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
      TabIndex        =   0
      ToolTipText     =   "Name"
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Local Error Resume Next
    Me.icon = frmMain.icon
    rtbDescription.BackColor = Color.Background
    rtbDescription.Font = BNET.Fonts
    rtbDescription.Font.size = BNET.FSize
    
    rtbLocation.BackColor = Color.Background
    rtbLocation.Font = BNET.Fonts
    rtbLocation.Font.size = BNET.FSize
    
    rtbUsername.BackColor = Color.Background
    rtbUsername.Font = BNET.Fonts
    rtbUsername.Font.size = BNET.FSize
    
    rtbSex.BackColor = Color.Background
    rtbSex.Font = BNET.Fonts
    rtbSex.Font.size = BNET.FSize
    
    rtbAge.BackColor = Color.Background
    rtbAge.Font = BNET.Fonts
    rtbAge.Font.size = BNET.FSize
End Sub

Public Sub SetProfile(strUser As String, sex As String, age As String, location As String, Description As String)
On Local Error Resume Next
    With sPB
        .InsertDWORD 1          'Number of accounts
        .InsertDWORD 4          'Number of items
        .InsertNTString strUser 'Accounts to update
        'Keys to update
        .InsertNTString "profile\sex"
        .InsertNTString "profile\age"
        .InsertNTString "profile\location"
        .InsertNTString "profile\description"
        'New values
        .InsertNTString sex
        .InsertNTString age
        .InsertNTString location
        .InsertNTString Description
        .SendPacket &H27
    End With
End Sub

Private Sub cmdClose_Click(): Unload Me: End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If frmMain.wsBNET.State = 7 And CloseProf Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub cmdSave_Click()
    SetProfile rtbUsername.text, rtbSex.text, rtbAge.text, rtbLocation.text, rtbDescription.text
    Me.Hide
End Sub

Private Sub rtbAge_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
    If KeyAscii = 13 Then
        If InStr(rtbAge.text, vbCrLf) = 0 Then rtbAge.text = rtbAge.text & vbCrLf
        rtbAge.SelStart = Len(rtbAge.text)
    End If
End Sub

Private Sub rtbLocation_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
    If KeyAscii = 13 Then
        If InStr(rtbLocation.text, vbCrLf) = 0 Then rtbLocation.text = rtbLocation.text & vbCrLf
        rtbLocation.SelStart = Len(rtbLocation.text)
    End If
End Sub

Private Sub rtbSex_KeyPress(KeyAscii As Integer)
On Local Error Resume Next
    If KeyAscii = 13 Then
        If InStr(rtbSex.text, vbCrLf) = 0 Then rtbSex.text = rtbSex.text & vbCrLf
        rtbSex.SelStart = Len(rtbAge.text)
        End If
End Sub

Private Sub rtbUsername_Change()
On Local Error Resume Next
    If LCase(rtbUsername.text) <> LCase(BNET.TrueUsername) Then
        Me.Caption = rtbUsername.text & "'s Profile"
        rtbAge.Locked = True
        rtbSex.Locked = True
        rtbLocation.Locked = True
        rtbDescription.Locked = True
        cmdSave.Enabled = False
    Else
        Me.Caption = "Your Profile"
        rtbAge.Locked = False
        rtbSex.Locked = False
        rtbLocation.Locked = False
        rtbDescription.Locked = False
        cmdSave.Enabled = True
    End If
End Sub

Public Sub AddText(ParamArray saElements() As Variant)
On Local Error Resume Next
    Dim i As Integer
    Dim data As String
    For i = LBound(saElements) To UBound(saElements) Step 2
        With saElements(i + 2)
            .SelStart = Len(.text)
            .SelLength = 0
            .SelColor = saElements(i)
            .SelText = saElements(i + 1) & Left$(vbCrLf, -2 * CLng((i + 1) = UBound(saElements)))
            .SelStart = Len(.text)
        End With
        data = data & saElements(i + 1)
    Next i
End Sub
