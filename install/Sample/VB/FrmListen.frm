VERSION 5.00
Begin VB.Form FrmListen 
   Caption         =   "Listen"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   Picture         =   "FrmListen.frx":0000
   ScaleHeight     =   2520
   ScaleWidth      =   4680
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check_Coach 
      Caption         =   "指导模式(客户听不到班长)"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "坐席列表"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "FrmListen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AgentList As String
Dim Agents As Collection

Private Sub CmdCancel_Click()
     
    If Me.Caption = "listen" Then
        frmMain.aOCX1.CmdListenStop
    Else
        frmMain.aOCX1.CmdIntrudeStop
    End If
    
    
    Unload Me
    
End Sub

Private Sub CmdOK_Click()
      If (Combo1.Text = "") Then
        MsgBox "Please select correct Agent"
        Exit Sub
    End If
    If Me.Caption = "listen" Then
        frmMain.aOCX1.CmdListenToAgent Combo1.Text
    Else
        If Check_Coach.Value = 1 Then
            '表示是Coach模式
            frmMain.aOCX1.CmdIntrudeAgent "$" & Combo1.Text
        Else
            frmMain.aOCX1.CmdIntrudeAgent Combo1.Text
        End If
        
        
    End If
    Unload Me
    
    
End Sub

Private Sub Form_Load()
Set Agents = New Collection
    
    Dim pos As Integer
    Dim posStart As Integer
    
    pos = 0
    posStart = 1
     
    pos = InStr(posStart, AgentList, ";")
    
    Dim tmpAgent As String
    
    Dim pos111 As Integer
    Dim tmpAgent1 As String
    
    While (pos >= 1)
        tmpAgent = Mid(AgentList, posStart, pos - posStart)
        pos111 = InStr(1, tmpAgent, "=")
        If (pos111 >= 1) Then
            tmpAgent1 = Left(tmpAgent, Len(tmpAgent) - 2)
        End If
        
        Agents.Add (tmpAgent1)
        posStart = pos + 1
        pos = InStr(pos + 1, AgentList, ";")
        
    Wend
    
    Dim i As Integer
    
    For i = 1 To Agents.Count
        Me.Combo1.AddItem Agents.Item(i)
    Next
    
    If Agents.Count > 0 Then
        Me.Combo1.ListIndex = 0
    Else
        Me.Combo1.Text = ""
    End If
    
    
    If Me.Caption = "listen" Then
        Check_Coach.Visible = False
    Else
        Check_Coach.Visible = True
        
    End If
    
    
    
End Sub
