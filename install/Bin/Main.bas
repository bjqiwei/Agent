'流程的进入点
'首先定位流程的第一个节点， 然后确定节点的性质， 调用相应的节点处理脚本
Sub OnLoad()
    Trace "Info:OnLoad()--- Main.bas"
    '首先定位xDoc中的第一个节点
    
  
    Dim NodesList
   
    Dim filter
    
    '首先取得第一个状态节点；必然是AgentRun下的第一个state节点
    filter = "//AgentRun/state"
    Set NodesList = xDoc.selectNodes(filter)
    
    '把CurrProcessNode定位为第一个state节点；
    Set CurrProcessNode = NodesList.Item(0)
    
    
    
    '跳转到第一个节点
    JumpToNode CurrProcessNode
    
    
    
End Sub



Sub JumpToNode(xNode)
	On Error Resume Next
	'根据结点类型进行跳转
	Select Case CurrProcessNode.baseName
	Case "state":
		'设置LastStateUuid：
    		LastStateUuid = CurrProcessNode.Attributes.getNamedItem("ID").nodeValue
    		
		NextModuleName =  "state"
		Exit Sub
	
	Case "Operation":
		NextModuleName =  "Operation"
		Exit Sub		
	Case Else:
		Trace "Err!子节点类型:" & CurrProcessNode.baseName & "无法处理！"
		
		Exit Sub	
	End Select
End Sub

Sub OnFrontEndEvent(EvtName)
On Error Resume Next
	Trace "在Main中捕捉到不合适FrontEnd事件：" & EvtName	
End Sub

Sub OnSoftPhoneEvent(EvtName)
On Error Resume Next
	Trace "在Main中捕捉到不合适SoftPhone命令：" & EvtName	
End Sub


Sub OnTimeOut()
	Trace "Info:收到不正确事件：TimeOut"
	
End Sub