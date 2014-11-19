VERSION 5.00
Begin VB.Form frmDialOut 
   Caption         =   "外拨"
   ClientHeight    =   1530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3705
   Icon            =   "frmDialOut.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3705
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txtPhoneNum 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "frmDialOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    If Me.Caption = "DialOut" Then
        frmMain.aOCX1.DialOut Trim(Me.txtPhoneNum.Text)
    ElseIf Me.Caption = "OutPhone" Then
        frmMain.aOCX1.OutPhone Trim(Me.txtPhoneNum.Text)
    ElseIf Me.Caption = "Conference" Then
        frmMain.aOCX1.Conference Trim(Me.txtPhoneNum.Text)
    End If
    
    DialOutMsg = Trim(Me.txtPhoneNum.Text)
    
    Unload Me
End Sub

Private Sub Command2_Click()
    If Me.Caption = "Conference" Then
        frmMain.aOCX1.CancelConference
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    Me.txtPhoneNum.Text = DialOutMsg
End Sub

Private Sub txtPhoneNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command1.SetFocus
    End If
End Sub
