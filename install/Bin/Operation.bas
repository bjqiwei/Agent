'操作节点
Sub OnLoad
	'先查找本节点对应的操作节点
	Trace "Info:OnLoad()--- Operation.bas"
	Dim filter
	Dim OperationNode, tmpNode
	filter = "./Operations"
	
	Set OperationNode = CurrProcessNode.SelectSingleNode(filter) '取得操作节点
	
	On Error Resume Next
	Dim i '用作计数
	
	If OperationNode Is Nothing Then '判断是否存在操作节点
		
	Else
	    Trace "Info:Yes Operation;Number is " & OperationNode.ChildNodes.Length
	    For i = 1 To OperationNode.ChildNodes.Length '有操作节点， 则对操作节点进行循环处理        
	    	
	    	Trace "Info:执行操作:" & OperationNode.ChildNodes(i-1).Attributes.getNamedItem("Expression").nodeValue
	    	Execute OperationNode.ChildNodes(i-1).Attributes.getNamedItem("Expression").nodeValue '进行操作
										    		
	    	
	    Next
	End If
	
	'执行完操作转入下一个节点
	filter = "./state"
	Set tmpNode = CurrProcessNode.selectSingleNode(filter)
	If tmpNode Is Nothing Then 
	    filter = "./Jump"
	    Set CurrProcessNode = CurrProcessNode.selectSingleNode(filter)
	Else    	
	    Set CurrProcessNode =  tmpNode
	End If
	JumpToNode CurrProcessNode
End Sub

Sub JumpToNode(xNode)
	On Error Resume Next
	'根据结点类型进行跳转
	Dim sName
	sName = CurrProcessNode.Attributes.getNamedItem("Name").nodeValue

	Select Case CurrProcessNode.baseName
	Case "state":
		If InStr(1,sName,"中") = 0 Then
			Trace "不是过渡状态"
			
		else
			Trace "是过渡状态"
		end if
		NextModuleName =  "state"
		Exit Sub
	
	Case "Operation":
		NextModuleName =  "Operation"
		Exit Sub		
	Case "Jump":
		NextModuleName = "Jump"
		Exit Sub
	Case Else:
		Trace "Info:Err!子节点类型:" & CurrProcessNode.baseName & "无法处理！"
		
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