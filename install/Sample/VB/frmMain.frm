VERSION 5.00
Object = "{88D896EC-5024-4605-A571-7E4B6C0CC8AD}#1.0#0"; "AgentOCX.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "座席"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10995
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   10995
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton IPOConf 
      Caption         =   "转自动"
      Height          =   375
      Left            =   8760
      TabIndex        =   36
      Top             =   2280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command9 
      Caption         =   "设置随路数据"
      Height          =   375
      Left            =   7440
      TabIndex        =   35
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "座席桌面监控"
      Height          =   375
      Left            =   6120
      TabIndex        =   34
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4800
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "呼叫历史"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   32
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "文本交谈"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   31
      Top             =   2280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "邮件处理"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   30
      Top             =   2280
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Caption         =   "自定义按钮"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   8400
      TabIndex        =   27
      Top             =   1845
      Width           =   1335
   End
   Begin VB.CheckBox Check2 
      Caption         =   "自动WrapEnd"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6600
      TabIndex        =   10
      Top             =   1845
      Width           =   1575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "自动摘机"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5400
      TabIndex        =   9
      Top             =   1845
      Width           =   1095
   End
   Begin VB.CommandButton cmdConfig 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   1440
      Picture         =   "frmMain.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   0
      Width           =   1215
   End
   Begin VB.CommandButton cmdLogout 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2880
      Picture         =   "frmMain.frx":2CA4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdLogin 
      Appearance      =   0  'Flat
      DragIcon        =   "frmMain.frx":5596
      Height          =   495
      Left            =   0
      Picture         =   "frmMain.frx":56E0
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.TextBox txtDNIS 
      Height          =   315
      Left            =   3960
      TabIndex        =   3
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtANI 
      Height          =   315
      Left            =   1320
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      Height          =   3765
      Left            =   0
      TabIndex        =   1
      Top             =   2760
      Width           =   10455
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   6600
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1766
            MinWidth        =   1766
            Text            =   "座席状态："
            TextSave        =   "座席状态："
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "座席ID："
            TextSave        =   "座席ID："
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3351
            MinWidth        =   3351
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4411
            MinWidth        =   4411
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   9600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      Picture         =   "frmMain.frx":7E82
      ScaleHeight     =   1095
      ScaleWidth      =   9615
      TabIndex        =   11
      Top             =   600
      Width           =   9615
      Begin VB.CommandButton Hold 
         Caption         =   "保持"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   25
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton Transfer 
         Caption         =   "转接"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   24
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton DialOut 
         Caption         =   "外拨"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   23
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton Consultation 
         Caption         =   "磋商"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   22
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton Auto 
         Caption         =   "磋商转接"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   21
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton OutPhone 
         Caption         =   "外线"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   20
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton Pause 
         Caption         =   "暂停"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Conference 
         Caption         =   "会议"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2640
         TabIndex        =   18
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Play 
         Caption         =   "放音"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3960
         TabIndex        =   17
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Listen 
         Caption         =   "监听"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   16
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Disconnect 
         Caption         =   "强拆"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6600
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton RopCall 
         Caption         =   "拦截"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Fax 
         Caption         =   "传真"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   13
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton Hook 
         Caption         =   "摘机"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   9615
      TabIndex        =   26
      Top             =   600
      Width           =   9615
      Begin AGENTOCXLibCtl.aOCX aOCX1 
         Height          =   975
         Left            =   0
         OleObjectBlob   =   "frmMain.frx":81C4
         TabIndex        =   28
         Top             =   0
         Width           =   9975
      End
   End
   Begin VB.Image ImageChat 
      Height          =   240
      Left            =   2880
      Picture         =   "frmMain.frx":81E8
      Top             =   2400
      Width           =   240
   End
   Begin VB.Image ImageMail 
      Height          =   240
      Left            =   1200
      Picture         =   "frmMain.frx":8332
      Top             =   2400
      Width           =   240
   End
   Begin VB.Label Label1 
      ForeColor       =   &H00008000&
      Height          =   255
      Left            =   4920
      TabIndex        =   29
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label DNIS 
      Caption         =   "被叫号码："
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   1845
      Width           =   975
   End
   Begin VB.Label ANI 
      Caption         =   "主叫号码："
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1845
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private Sub ActiveBar21_ToolClick(ByVal Tool As ActiveBar2LibraryCtl.Tool)
'    Select Case Tool.Name
'        Case "tlLogin"
'            frmLogin.Show
'        Case "tlLogoff"
'            aOCX1.LogOUT
'            Me.ActiveBar21.Bands("Band1").Tools("tlLogin").Enabled = True
'        Case "tlConfig"
'            Me.aOCX1.ShowConfig
'    End Select
'End Sub
Public DAgentID As String
Public gAgentID As String

Dim g_Status As String
Dim AgentIDSpeaked As Boolean



Private Sub aOCX1_BBSCallArrive(ByVal MessageID As Long)
    'MsgBox "BBSCall " & MessageID
End Sub

Private Sub aOCX1_CallArrive(ByVal ANI As String, ByVal DNIS As String, ByVal data As String)
    AgentIDSpeaked = False
    txtANI.Text = ""
    txtDNIS.Text = ""
    
    txtANI.Text = ANI
    txtDNIS.Text = DNIS
    'MsgBox "CallArrive"
    
    If data <> "" Then
        'MsgBox data
    End If
    
    If Check1.Value = 1 Then
        aOCX1.CmdAnswer
    End If
    
End Sub

Private Sub aOCX1_DialTaskArrive(ByVal DialerID As String)
    'MsgBox "这是外拨任务,DialerID是:" & DialerID
End Sub

Private Sub aOCX1_EMailCallArrive(ByVal Count As Long)
    If (Count > 0) Then
        ImageMail.Visible = True
    Else
        ImageMail.Visible = False
    End If
    'Label2.Caption = Count
    'Label2.Visible = True
End Sub

Private Sub aOCX1_EVTAgentAvail()
    'Label1.Visible = False
   
        Label1.Visible = True
        Label1.ForeColor = &H8000&
        Label1.Caption = "硬件电话状态闲，可以接受呼叫，无法外拨"
   
End Sub

Private Sub aOCX1_EVTAgentOther()
    'Label1.Visible = True
   
        Label1.Visible = True
        Label1.ForeColor = &HFF&
        Label1.Caption = "硬件电话状态忙，无法接受呼叫，可以外拨"
   
End Sub

Private Sub aOCX1_EVTAnswerSucc()
    If AgentIDSpeaked = False Then
        
        
        'Me.aOCX1.CmdPlayAgentIDWelcome
        AgentIDSpeaked = True
    End If
    
End Sub

Private Sub aOCX1_EVTButtonStatus(ByVal Name As String, ByVal Title As String, ByVal Enable As Long)
    Debug.Print Name, Title, Enable
    
    Select Case Title
        Case "OnHook"
            Title = "挂机"
        Case "OffHook"
            Title = "摘机"
        Case "HOOKOFFCONSULT"
            Title = "接受"
        Case "Hold"
            Title = "保持"
        Case "HoldCancel"
            Title = "取消"
        Case "Transfer"
            Title = "转接"
        Case "CancelTransfer"
            Title = "取消"
        Case "DialOut"
            Title = "外拨"
        Case "CancelDialOut"
            Title = "取消"
        Case "Consultation"
            Title = "磋商"
        Case "CancelConsultation"
            Title = "取消"
        Case "StopConsultation"
            Title = "磋商结束"
        Case "ConsultTransfer"
            Title = "磋商转接"
        Case "Auto"
            Title = "磋商转接"
        Case "OutPhone"
            Title = "自动"
        Case "CancelOutPhone"
            Title = "取消"
        Case "Play"
            Title = "放音"
        Case "PlayCancel"
            Title = "结束"
        Case "Fax"
            Title = "传真"
        Case "FaxStop"
            Title = "结束"
        Case "Pause"
            Title = "暂停"
        Case "Continue"
            Title = "恢复"
        Case "ContinueDialTask"
            Title = "放弃回访"
        Case "Listen"
            Title = "监听"
        Case "CancelListen"
            Title = "结束"
        Case "Disconnect"
            Title = "强插"
        Case "Conference"
            Title = "会议"
        Case "CancelConference"
            Title = "取消"
        Case "RopCall"
            Title = "拦截"
        Case "LOGINSUCC"
            Title = "登录成功"
        Case "LOGINFAIL"
            Title = "登录失败"
        Case "TRANSFERFAIL"
            Title = "转接失败"
    End Select
    
    Select Case Name
        Case "Hook"
            SetButtonOption Title, Enable, Me.Hook
        Case "Hold"
            SetButtonOption Title, Enable, Me.Hold
        Case "Transfer"
            SetButtonOption Title, Enable, Me.Transfer
        Case "DialOut"
            SetButtonOption Title, Enable, Me.DialOut
        Case "Consultation"
            SetButtonOption Title, Enable, Me.Consultation
        Case "Auto"
            SetButtonOption Title, Enable, Me.Auto
        Case "OutPhone"
            SetButtonOption Title, Enable, Me.OutPhone
        Case "Fax"
            SetButtonOption Title, Enable, Me.Fax
        Case "Pause"
            SetButtonOption Title, Enable, Me.Pause
           
        Case "Conference"
            SetButtonOption Title, Enable, Me.Conference
        Case "Play"
            SetButtonOption Title, Enable, Me.Play
        Case "Listen"
            SetButtonOption Title, Enable, Me.Listen
        Case "Disconnect"
            SetButtonOption Title, Enable, Me.Disconnect
        Case "RopCall"
            SetButtonOption Title, Enable, Me.RopCall
        Case "Init"
            SetButtonInit
        'Case Else
         '   SetButtonInit
    End Select
    
End Sub

Private Sub aOCX1_EVTConference()

    frmDialOut.Caption = "Conference"
    frmDialOut.Show vbModal
End Sub

Private Sub aOCX1_EVTConsult(ByVal AgentList As String)
    'AgentList ==> Agent1;Agent2;...
    'Form1.Text1.Text = AgentList
    
    'Form1.Show vbModal
    FrmConsult.AgentList = AgentList
    FrmConsult.Show 1
End Sub

Private Sub aOCX1_EVTConsultSucc(ByVal sAgentID As String)
'    Me.cmdConsultTransfer.Enabled = True
'    Me.CmdConsultStop.Enabled = True
    DAgentID = sAgentID
End Sub

Private Sub aOCX1_EVTDialOut()
    
    frmDialOut.Caption = "DialOut"
    frmDialOut.Show vbModal
End Sub

Private Sub aOCX1_EVTFax()
    frmFax.Caption = "FaxOut"
    frmFax.Show vbModal
End Sub

Private Sub aOCX1_EVTFreeAgentsList(ByVal AgentList As String)
    'MsgBox AgentList
End Sub

Private Sub aOCX1_EVTIntrude(ByVal AgentList As String)
'    FrmListen.AgentList = AgentList
'    FrmListen.Caption = "intrude"
'
'    FrmListen.Show 1
End Sub

Private Sub aOCX1_EVTIntrude2(ByVal AgentList As String)
      On Error GoTo errHandle

    'MyLog "aOCX1_EVTIntrude2 begin"
    
    FrmListen2.AgentList = AgentList
    Set FrmListen2.aOCX1 = aOCX1
    
    FrmListen2.Caption = "intrude"
    
    FrmListen2.Show 1
    
    'MyLog "aOCX1_EVTIntrude2 end"
    
    'PlsWait
    Exit Sub
errHandle:
    MsgBox Err.Number & ":" & Err.Description
'    localfile "aOCX1_EVTIntrude2 :" & Err.Number & ":" & Err.Description
End Sub

Private Sub aOCX1_EVTListen(ByVal AgentList As String)
'    FrmListen.AgentList = AgentList
'    FrmListen.Caption = "listen"
'
'    FrmListen.Show 1
End Sub

Private Sub aOCX1_EVTListen2(ByVal AgentList As String)
     On Error GoTo errHandle

    'MyLog "aOCX1_EVTListen2 begin"
    
    FrmListen2.AgentList = AgentList
    Set FrmListen2.aOCX1 = aOCX1
    
    FrmListen2.Caption = "listen"
    
    FrmListen2.Show 1
    
    'MyLog "aOCX1_EVTListen2 end"
    
    'PlsWait
    Exit Sub
errHandle:
    MsgBox Err.Number & ":" & Err.Description
    'localfile "aOCX1_EVTListen2 :" & Err.Number & ":" & Err.Description    '
End Sub

Private Sub aOCX1_EVTLoginFailed(ByVal reason As String)
    MsgBox reason
End Sub

Private Sub aOCX1_EVTLoginSuc()
    'Me.ActiveBar21.Bands("Band1").Tools("tlLogin").Enabled = False
    Me.cmdLogin.Enabled = False
    aOCX1.TextChatLogin
    aOCX1.MisStatus = "我在吃饭"
   
    
    
    'MsgBox "LoginSucc"
    
End Sub

Private Sub aOCX1_EVTMakeCallFailed()
    MsgBox "EVTMakeCallFailed"
End Sub

Private Sub aOCX1_EVTMakeCallFailedByReason(ByVal reason As String)
    MsgBox reason
    
End Sub

Private Sub aOCX1_EVTOutPhone()

    frmDialOut.Caption = "OutPhone"
    frmDialOut.Show vbModal
End Sub


Private Sub aOCX1_EVTReturnStatus(ByVal rstatus As String)
    g_Status = rstatus
    Debug.Print rstatus
    Me.StatusBar1.Panels(2).Text = rstatus
    Me.StatusBar1.Refresh
     
    If Len(rstatus) > 40 Then
        'SetButtons rstatus
        
    Else
        If rstatus = "连接断开" Then
            'Me.ActiveBar21.Bands("Band1").Tools("tlLogin").Enabled = True
            Me.cmdLogin.Enabled = True
        End If
        
        If rstatus <> "空闲状态" Then
            Label1.Visible = False
        End If
        
    End If
        
End Sub


Private Sub aOCX1_EVTReturnStatusCH(ByVal Status As String)
    Debug.Print Status
    
End Sub

Private Sub aOCX1_EVTTimeOut(ByVal Status As String)
    'MsgBox Status & "-->超时"
End Sub

Private Sub aOCX1_EVTTransfer(ByVal AgentList As String)
    'MsgBox AgentList
    'If Left(Me.Transfer.Caption, 2) = "转接" Then
    'Me.aOCX1.CmdTransfer
    'Else
    '    Me.aOCX1.CmdTransferStop
    'End If
    
    FrmTransfer.AgentList = AgentList
    FrmTransfer.Show 1
End Sub

Private Sub aOCX1_EVTWrapUp()
    If Me.Check2.Value = 1 Then
        Me.aOCX1.WrapEnd
    End If
End Sub

Private Sub aOCX1_OACallArrive(ByVal MessageID As Long)
    'MsgBox "OACall " & MessageID
End Sub

Private Sub aOCX1_RunWordChange(ByVal sRunWord As String)
    MsgBox sRunWord
    
End Sub

Private Sub aOCX1_TextChatLoginFail(ByVal reason As String)
    'MsgBox "TextChatLoginFail..." & reason
    ImageChat.Visible = False
End Sub

Private Sub aOCX1_TextChatLoginSucc()
    'MsgBox "TextChatLoginSucc"
    ImageChat.Visible = True
End Sub

Private Sub aOCX1_TextChatLogOuted()
    ImageChat.Visible = False
End Sub

Private Sub Auto_Click()
    Me.aOCX1.CmdConsultTransfer DAgentID, "", "", ""
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Me.Picture1.Visible = True
        Me.Picture2.Visible = False
    Else
        Me.Picture1.Visible = False
        Me.Picture2.Visible = True
    End If
End Sub

Private Sub cmdConfig_Click()
    Me.aOCX1.ShowConfig
End Sub

Private Sub cmdConsultStop_Click()
'    Me.aOCX1.CmdConsultStop
'    Me.CmdConsultStop.Enabled = False
End Sub

Private Sub cmdConsultTransfer_Click()
    'Me.aOCX1.CmdConsultToAgent
'    Me.cmdConsultTransfer.Enabled = False
    
End Sub

Private Sub cmdLogin_Click()
    frmLogin.Show
End Sub

Private Sub cmdLogout_Click()
    aOCX1.LogOUT
    Me.cmdLogin.Enabled = True
    
    SetButtonInit
End Sub

Private Sub cMSN_Click()
'    If cMSN.Value = 1 Then
'        msnX1.Start
'    Else
'        msnX1.Bye
'
'    End If
End Sub

Private Sub Command1_Click()
    'FrmConsult.Show 1
    Me.aOCX1.GoMail
    Exit Sub
End Sub

Private Sub Command2_Click()
    'aOCX1.TestBarTip "zhutong"
    'aOCX1.TextChatLogin
    aOCX1.ShowCallHistory
End Sub

Private Sub Command3_Click()
    aOCX1.ShowTextChatDlg
End Sub

Private Sub Command4_Click()
    'aOCX1.CmdQueryFreeAgentsList
    'aOCX1.StartWatchDeskTop "192.168.1.107"
    aOCX1.QueryAgentAddresses
End Sub

Private Sub Command5_Click()
    FrmListen.AgentList = "Agent1=0;"
    FrmListen.Show 1
End Sub

Private Sub Commandmsn_Click()
    'msnX1.Start
End Sub

Private Sub Command9_Click()
    'aOCX1.DoSetAssociatedData "AGENTOCX", "你好"
    MsgBox aOCX1.DoGetAssociatedData("TTS")
End Sub

Private Sub Conference_Click()
    If Left(Me.Conference.Caption, 2) = "会议" Then
        Me.aOCX1.CmdConference
    Else
        Me.aOCX1.CmdMakeCallStop
    End If
End Sub

Private Sub Consultation_Click()
    If Left(Me.Consultation.Caption, 4) = "磋商结束" Then
        Me.aOCX1.CmdConsultStop
    ElseIf Left(Me.Consultation.Caption, 2) = "磋商" Then
        Me.aOCX1.CmdConsult
        'FrmConsult.Show 1
        
    Else
        Me.aOCX1.CmdConsultStop
    End If
End Sub

Private Sub DialOut_Click()
    If Left(Me.DialOut.Caption, 2) = "外拨" Then
        Me.aOCX1.CmdDialOut
    Else
        Me.aOCX1.CmdMakeCallStop
        
    End If
End Sub


Private Sub Disconnect_Click()
    Me.aOCX1.CmdIntrude
End Sub

Private Sub Fax_Click()
    Me.aOCX1.CmdFax
End Sub

Private Sub Form_Load()
    
    
    ImageChat.Visible = False
    ws.Protocol = sckUDPProtocol
    ws.Bind 33334
    frmMain.aOCX1.PlayFileName = "c:\play.vox"
    Me.txtANI.Text = ""
    Me.txtDNIS.Text = ""
    ImageMail.Visible = False
    
'    Dim saShortcuts1(1) As New ShortCut
'    saShortcuts1(1).Value = "F1"
'    Me.ActiveBar21.Bands("Band1").Tools("tlLogin").ShortCuts = saShortcuts1
'
'    Dim j
'    j = Me.ActiveBar21.Bands("Band1").Tools("tlLogin").ShortCuts(1).Value
'
'    j = saShortcuts1(1).Value

    Hook.Enabled = False
    Hold.Enabled = False
    Transfer.Enabled = False
    DialOut.Enabled = False
    Consultation.Enabled = False
    Auto.Enabled = False
    OutPhone.Enabled = False
    Fax.Enabled = False
    Pause.Enabled = False
    Conference.Enabled = False
    Play.Enabled = False
    Listen.Enabled = False
    Disconnect.Enabled = False
    RopCall.Enabled = False
    'Label2.Visible = False
'    Me.cmdConsultTransfer.Enabled = False
'    Me.CmdConsultStop.Enabled = False
                
     aOCX1.PicSize = 320
     
     Check3_Click
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
'    Me.aOCX1.LogOUT
End Sub

Private Sub Hold_Click()
    If Left(Me.Hold.Caption, 2) = "保持" Then
        Me.aOCX1.CmdHold
    Else
        Me.aOCX1.CmdHoldStop
    End If
End Sub

Private Sub Hook_Click()
    If Left(Me.Hook.Caption, 2) = "摘机" Then
        Me.aOCX1.CmdAnswer
        
        
    ElseIf Me.Hook.Caption = "接受" Then
        Me.aOCX1.CmdConsultAnswer
    Else
        Me.aOCX1.DoSetAssociatedData "AGENTID", Me.gAgentID
        'Me.aOCX1.cmdAuto2 "ToAuto", "BLINDTRANSFER"
        
        
        Me.aOCX1.CmdOnHook
    End If
End Sub

Private Sub ImageMail_Click()
    'aOCX1.GoMail
End Sub

Private Sub IPOConf_Click()
    'frmIPOConf.Show 1
    'aOCX1.CmdSetRunWords
    
    'aOCX1.cmdAuto2 "ToAuto", ""
    
    aOCX1.CmdTransferToAgent "Agent3", "", "", ""
    
    
End Sub

Private Sub Listen_Click()
    If Left(Me.Listen.Caption, 2) = "监听" Then
        Me.aOCX1.CmdListen
    Else
        Me.aOCX1.CmdListenStop
    End If
    
End Sub

Private Sub OutPhone_Click()
    
'    If Left(Me.OutPhone.Caption, 2) = "外线" Then
'        Me.aOCX1.CmdOutPhone
'    Else
'        Me.aOCX1.CmdMakeCallStop
'    End If
    Me.aOCX1.CmdAuto
    
End Sub

Private Sub Pause_Click()
    
    If Left(Me.Pause.Caption, 2) = "暂停" Then
        Me.aOCX1.CmdPause
    Else
        Me.aOCX1.CmdContinue
    End If
End Sub

Private Sub Play_Click()
    If Left(Me.Play.Caption, 2) = "放音" Then
        Me.aOCX1.CmdPlay
    Else
        Me.aOCX1.CmdPlayStop
    End If
End Sub

Private Sub RopCall_Click()
    Me.aOCX1.CmdRopCall
End Sub

Private Sub Timer1_Timer()
    
End Sub

Private Sub Transfer_Click()
    If Left(Me.Transfer.Caption, 2) = "转接" Then
        Me.aOCX1.CmdTransfer
        
    Else
        Me.aOCX1.CmdTransferStop
    End If
End Sub

Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    Dim sdata As String
    ws.GetData sdata
    
    List1.AddItem sdata
    If Not List1.ListCount = 0 Then
        List1.ListIndex = List1.ListCount - 1
    End If
End Sub


