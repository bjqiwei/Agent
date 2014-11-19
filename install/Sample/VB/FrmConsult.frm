VERSION 5.00
Begin VB.Form FrmConsult 
   Caption         =   "Consult"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   Picture         =   "FrmConsult.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancel 
      Caption         =   "取消"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "确定"
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   360
      Width           =   1815
   End
   Begin VB.OptionButton Option2 
      Caption         =   "座席磋商："
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "外线磋商："
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "FrmConsult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public AgentList As String
Dim Agents As Collection

Private Sub CmdCancel_Click()
        
    frmMain.aOCX1.CmdConsultCancel
    
    Unload Me
End Sub

Private Sub CmdOK_Click()
    
    If Option1.Value = True Then
        If Trim(Text1.Text) = "" Then
            MsgBox "请输入正确的电话号码！"
            Exit Sub
        End If
        frmMain.aOCX1.CmdConsultToOutLine Me.Text1.Text
    Else
        If Trim(Combo1.Text) = "" Then
            MsgBox "请选择正确的座席！"
            Exit Sub
        End If
        frmMain.aOCX1.CmdConsultToAgent Me.Combo1.Text, "", ""
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    '解析AgentList
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
    Me.Text1.Text = ""
End Sub

