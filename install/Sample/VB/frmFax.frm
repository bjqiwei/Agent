VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmFax 
   Caption         =   "FaxOut"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5355
   Icon            =   "frmFax.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2145
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdButton 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   6
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox txtFaxName 
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      Text            =   "E:\1.tif"
      Top             =   840
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox txtPhoneNum 
      Height          =   315
      Left            =   1200
      TabIndex        =   1
      Text            =   "103"
      Top             =   300
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "传真文件："
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "电话号码："
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "frmFax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdButton_Click()
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
  
    CommonDialog1.Filter = "All Files (*.*)|*.*|Tiff Files (*.tif)|*.tif|Text Files" & _
    "(*.txt)|*.txt"
  
    Me.CommonDialog1.ShowOpen
    Me.txtFaxName.Text = Me.CommonDialog1.FileName
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub
End Sub

Private Sub Command1_Click()
    frmMain.aOCX1.FaxSend Trim(Me.txtPhoneNum.Text), Trim(Me.txtFaxName.Text)
    FaxOutMsg = Trim(Me.txtPhoneNum.Text)
    FaxOutFileMsg = Trim(Me.txtFaxName.Text)
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If FaxOutFileMsg = "" Then
        Me.txtFaxName.Text = "e:\1.tif"
    Else
        Me.txtFaxName.Text = FaxOutFileMsg
    End If
    If FaxOutMsg = "" Then
        Me.txtPhoneNum.Text = ""
    Else
        Me.txtPhoneNum.Text = FaxOutMsg
    End If
End Sub

Private Sub txtFaxName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Command1.SetFocus
    End If
End Sub

Private Sub txtPhoneNum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFaxName.SetFocus
    End If
End Sub
