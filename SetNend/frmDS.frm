VERSION 5.00
Begin VB.Form frmDS 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Death Spam"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   Icon            =   "frmDS.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDS.frx":0442
   ScaleHeight     =   3345
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTimes 
      Height          =   330
      Left            =   4455
      TabIndex        =   10
      Text            =   "1"
      Top             =   135
      Width           =   600
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send"
      Height          =   330
      Left            =   3600
      MousePointer    =   2  'Cross
      TabIndex        =   5
      Top             =   540
      Width           =   1500
   End
   Begin VB.ListBox lstMsgs 
      Height          =   2010
      Left            =   45
      TabIndex        =   7
      Top             =   1350
      Width           =   5055
   End
   Begin VB.TextBox txtFrom 
      Height          =   330
      Left            =   990
      TabIndex        =   1
      Top             =   540
      Width           =   2535
   End
   Begin VB.TextBox txtTo 
      Height          =   330
      Left            =   990
      TabIndex        =   0
      Top             =   135
      Width           =   2535
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4365
      MousePointer    =   2  'Cross
      TabIndex        =   4
      Top             =   945
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3600
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   945
      Width           =   735
   End
   Begin VB.TextBox txtMsg 
      Height          =   330
      Left            =   990
      TabIndex        =   2
      Top             =   945
      Width           =   2535
   End
   Begin VB.Label lblTimes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Times:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3735
      TabIndex        =   11
      Top             =   270
      Width           =   645
   End
   Begin VB.Label lblFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   375
      TabIndex        =   9
      Top             =   675
      Width           =   570
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   45
      TabIndex        =   8
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   645
      TabIndex        =   6
      Top             =   270
      Width           =   300
   End
End
Attribute VB_Name = "frmDS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SendTimes As Integer
Dim from As String
Dim sendto As String
Dim a As Integer
Dim b As Integer
Dim res As Boolean

Private Sub cmdAdd_Click()
If Trim(txtMsg.Text) <> "" Then
    lstMsgs.AddItem txtMsg.Text
    txtMsg.Text = ""
Else
    MsgBox "Enter message first!", vbCritical + vbDefaultButton1, "DS"
    txtMsg.SetFocus
End If

End Sub

Private Sub cmdDel_Click()
If lstMsgs.ListIndex <> -1 Then
    lstMsgs.RemoveItem lstMsgs.ListIndex
End If
End Sub

Private Sub cmdSend_Click()
If Trim(txtTo.Text) = "" Then
    MsgBox "Enter user/computer to send to!", vbCritical + vbDefaultButton1, "DS"
    txtTo.SetFocus
    Exit Sub
End If

If CInt(txtTimes.Text) >= 100 Then
    MsgBox "Easy on the spamming, 99 is the maximum - for obvious reasons", vbCritical + vbDefaultButton1, "DS"
    txtTimes.SetFocus
    Exit Sub
End If

If Trim(txtFrom.Text) = "" Then
    MsgBox "Enter sender name!", vbCritical + vbDefaultButton1, "DS"
    txtFrom.SetFocus
    Exit Sub
End If

If lstMsgs.ListCount < 1 Then
    MsgBox "Enter some messages!", vbCritical + vbDefaultButton1, "DS"
    txtMsg.SetFocus
    Exit Sub
End If

sendto = txtTo.Text
from = txtFrom.Text
SendTimes = CInt(txtTimes.Text)
If SendTimes = 0 Then SendTimes = 1

For a = 1 To SendTimes
    For b = 0 To lstMsgs.ListCount
        res = frmMain.BroadcastMessage(sendto, from, lstMsgs.List(b))
        If res = False Then
            MsgBox "Death spam failed on message " & b & " of " & lstMsgs.ListCount, vbCritical + vbDefaultButton1, "DS"
            Exit Sub
        End If
    Next b
Next a
If res = True Then
    MsgBox "Death Spam was successful!", vbInformation + vbDefaultButton1, "DS"
End If
End Sub

Private Sub Form_Load()
txtFrom.Text = Environ$("USERNAME")
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub
