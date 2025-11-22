VERSION 5.00
Begin VB.Form frmD2DV 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create a Diablo 2 Game"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmCrGameD2 
      Caption         =   "Create Game"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
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
         TabIndex        =   16
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton btmCrGameD2 
         Caption         =   "&Create Game"
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
         Left            =   3000
         TabIndex        =   6
         Top             =   3840
         Width           =   1335
      End
      Begin VB.OptionButton optNightmare 
         Caption         =   "Nightmare"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         Top             =   3480
         Width           =   1695
      End
      Begin VB.OptionButton optNormal 
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   3480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.CheckBox chkPDiffD2 
         Caption         =   "Level restrictions +/-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox txtPDiffD2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   5
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txtMaxPD2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         MaxLength       =   1
         TabIndex        =   4
         Text            =   "4"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtGDescD2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaxLength       =   31
         TabIndex        =   3
         Top             =   2040
         Width           =   3975
      End
      Begin VB.TextBox txtGPassD2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtGNameD2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         MaxLength       =   15
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
      Begin VB.OptionButton optHell 
         Caption         =   "Hell"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   15
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "levels"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   2520
         Width           =   1080
      End
      Begin VB.Label lblMaxPD2 
         BackStyle       =   0  'Transparent
         Caption         =   "Max. number of players:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label lblGDescD2 
         Caption         =   "Game Description"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lblPasswordD2 
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
         Left            =   240
         TabIndex        =   8
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label lblGNameD2 
         Caption         =   "Game Name"
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
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmD2DV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btmCrGameD2_Click()
Dim d2Diff As String, d2PDiff As String
    d2PDiff = txtPDiffD2.text
    If optNormal.value = True Then d2Diff = 0
    If optNightmare.value = True Then d2Diff = 1
    If optHell.value = True Then d2Diff = 2
    If Not chkPDiffD2.Enabled Then d2PDiff = "FF"
    Call sndCrGame(d2Diff, d2PDiff, txtMaxPD2.text, txtGNameD2.text, txtGPassD2.text, txtGDescD2.text)
End Sub

Private Sub chkPDiffD2_Click()
    If chkPDiffD2.Enabled Then
        txtPDiffD2.Enabled = True
    Else
        txtPDiffD2.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub optNormal_Click()
    If optNormal.value = True Then
        optNightmare.value = False
        optHell.value = False
    End If
End Sub

Private Sub optNightmare_Click()
    If optNightmare.value = True Then
        optNormal.value = False
        optHell.value = False
    End If
End Sub

Private Sub optHell_Click()
    If optHell.value = True Then
        optNormal.value = False
        optNightmare.value = False
    End If
End Sub

Private Sub txtMaxPD2_Change()
    If CInt(txtMaxPD2.text) > 8 Then txtMaxPD2.text = "8"
End Sub
