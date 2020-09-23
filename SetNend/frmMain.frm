VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Set Nend"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0442
   ScaleHeight     =   3825
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDS 
      BackColor       =   &H00000000&
      Caption         =   "&Death Spam"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2205
      MaskColor       =   &H00E0E0E0&
      MousePointer    =   2  'Cross
      TabIndex        =   7
      Top             =   900
      UseMaskColor    =   -1  'True
      Width           =   1275
   End
   Begin VB.CommandButton cmdSend 
      BackColor       =   &H00000000&
      Caption         =   "&Send"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3510
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   900
      Width           =   1275
   End
   Begin VB.TextBox txtMsg 
      Height          =   2040
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1755
      Width           =   4740
   End
   Begin VB.TextBox txtTo 
      Height          =   330
      Left            =   1980
      TabIndex        =   0
      Top             =   90
      Width           =   2805
   End
   Begin VB.TextBox txtFrom 
      Height          =   330
      Left            =   1980
      TabIndex        =   1
      Top             =   495
      Width           =   2805
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
      Left            =   1335
      TabIndex        =   6
      Top             =   630
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
      Left            =   90
      TabIndex        =   5
      Top             =   1485
      Width           =   900
   End
   Begin VB.Label lblTo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To user/computer:"
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
      Left            =   60
      TabIndex        =   4
      Top             =   225
      Width           =   1845
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function NetMessageBufferSend Lib "NETAPI32.DLL" _
(yServer As Any, yToName As Byte, yFromName As Any, yMsg As Byte, _
ByVal lSize As Long) As Long
Private Const NERR_Success As Long = 0&

Public Function BroadcastMessage(UserOrMachine As String, _
FromName2 As String, Message As String) As Boolean
   
    Dim ToName() As Byte
    Dim FromName() As Byte
    Dim MessageToSend() As Byte
    
    'Put data into byte arrays
    ToName = UserOrMachine & vbNullChar
    FromName = FromName2 & vbNullChar
    MessageToSend = Message & vbNullChar
    
    'Broadcast message via API
    If NetMessageBufferSend(ByVal 0&, ToName(0), FromName(0), MessageToSend(0), UBound(MessageToSend)) = NERR_Success Then
        'Return True if it worked
        BroadcastMessage = True
    End If

End Function

Private Sub cmdDS_Click()
frmDS.Show
End Sub

Private Sub cmdSend_Click()
If Validate = True Then
    If BroadcastMessage(txtTo.Text, txtFrom.Text, txtMsg.Text) Then
        MsgBox "Message sent!", vbInformation + vbDefaultButton1, "Set Nend"
    Else
        MsgBox "Message failed to send.", vbCritical + vbDefaultButton1, "Set Nend"
    End If
End If
End Sub

Private Function Validate() As Boolean
If Len(Trim(txtTo.Text)) < 1 Then
    MsgBox "Please enter user/computer to send to.", vbInformation
    txtTo.SetFocus
    Validate = False
    Exit Function
End If
If Len(Trim(txtFrom.Text)) < 1 Then
    MsgBox "Please enter who it is from.", vbInformation
    txtFrom.SetFocus
    Validate = False
    Exit Function
End If
If Len(Trim(txtMsg.Text)) < 1 Then
    MsgBox "Please enter a message", vbInformation
    txtMsg.SetFocus
    Validate = False
    Exit Function
End If
Validate = True
End Function

Private Sub Form_Load()
txtFrom.Text = Environ$("USERNAME")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmDS
End Sub
