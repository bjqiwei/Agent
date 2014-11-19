VERSION 5.00
Begin VB.Form frmIPOConf 
   Caption         =   "指定IPO会议码"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3600
   LinkTopic       =   "Form2"
   Picture         =   "frmIPOConf.frx":0000
   ScaleHeight     =   1380
   ScaleWidth      =   3600
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.TextBox txtPhoneNum 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   2160
      TabIndex        =   0
      Top             =   840
      Width           =   1335
   End
End
Attribute VB_Name = "frmIPOConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If Len(Me.txtPhoneNum) = 0 Then
        MsgBox "请输入会议代码"
        Exit Sub
    End If
    
    frmMain.aOCX1.CmdConf_IPO Me.txtPhoneNum
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub
