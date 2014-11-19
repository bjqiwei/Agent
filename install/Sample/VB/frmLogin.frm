VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "登录"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4335
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2700
   ScaleWidth      =   4335
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   30
      Left            =   0
      TabIndex        =   7
      Top             =   1920
      Width           =   4340
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      Picture         =   "frmLogin.frx":2072
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   6
      Top             =   360
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "取消"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfirm 
      BackColor       =   &H00C0C0FF&
      Caption         =   "确定"
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Text            =   "2"
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   2040
      TabIndex        =   1
      Text            =   "Agent2"
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "密码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "用户名："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      TabIndex        =   0
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdConfirm_Click()
    'If Trim(Me.txtUserName.Text) = "" Or Trim(Me.txtPassword.Text) = "" Then
     '   MsgBox "座席ID和密码不许为空！", vbOKOnly, "座席"
    'Else
    
    
    Dim cReg As New cRegistry
    With cReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Samwoo\AA\PBXSoftPhone"

        .ValueKey = "AgentID"
        .ValueType = REG_SZ
        
        AgentID = Me.txtUserName
        
        .Value = AgentID
        
        
    End With
     
    If True Then
        On Error GoTo JumpErr
        'frmMain.aOCX1.PicSize = -1
        frmMain.aOCX1.AgentID = Trim(Me.txtUserName.Text)
        frmMain.aOCX1.Password = Trim(Me.txtPassword.Text)
        frmMain.aOCX1.CTIServerIP = "192.168.1.103"
        frmMain.aOCX1.ShowShortCut = 0
        frmMain.aOCX1.AutoAnswer = 0
        frmMain.gAgentID = Trim(Me.txtUserName.Text)
        
        
        If frmMain.Check3.Value = 0 Then
            frmMain.aOCX1.SetToolsVisible 1
            frmMain.Picture1.Visible = False
            frmMain.Picture2.Visible = True
        Else
            frmMain.aOCX1.SetToolsVisible 1
            frmMain.Picture1.Visible = True
            frmMain.Picture2.Visible = False
        End If
                
        'MsgBox frmMain.aOCX1.AutoAnswer
        frmMain.aOCX1.LogIn
        LoginMsg = Trim(Me.txtUserName.Text)
        LoginPMsg = Trim(Me.txtPassword.Text)
        frmMain.StatusBar1.Panels(4).Text = Trim(Me.txtUserName.Text)
        
        Unload Me
        Exit Sub
JumpErr:
        frmMain.StatusBar1.Panels(4).Text = ""
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    If LoginMsg = "" Then
        LoginMsg = "Agent2"
    End If
    If LoginPMsg = "" Then
        LoginPMsg = "2"
    End If
    
    Me.txtUserName = LoginMsg
    Me.txtPassword = LoginPMsg
    
    
    Dim cReg As New cRegistry
    With cReg
        .ClassKey = HKEY_LOCAL_MACHINE
        .SectionKey = "Software\Samwoo\AA\PBXSoftPhone"

        .ValueKey = "AgentID"
        .ValueType = REG_SZ
        AgentID = .Value
        
        If .Value = Null Or .Value = "" Then
            .Value = "Agent2"
            AgentID = .Value
        End If
        
     End With
     
     Me.txtUserName = AgentID
     
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdConfirm.SetFocus
    End If
End Sub

Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtPassword.SetFocus
    End If
End Sub
