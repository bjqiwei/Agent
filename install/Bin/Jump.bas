'Jump节点

Sub OnLoad()
	'找到Destination节点
	
	 Trace "Info:OnLoad()--- Jump.bas"
	Dim NodeList
	Dim destination
	Dim xNodeTmp
	Dim filter
	Dim bFound 
	
	bFound = False
	
	'先取出Destination属性
	destination = CurrProcessNode.attributes.getNamedItem("Destination").nodeValue
	
	filter = "//*[@ID=""" & destination & """]"
	Set xNodeTmp = xDoc.selectSingleNode(filter)
	If Not xNodeTmp Is Nothing Then
			
		Set CurrProcessNode = xNodeTmp
		bFound = True
				
	End If
	
	If bFound = True Then
		
		
		JumpToNode CurrProcessNode
		Exit Sub
	Else
		TRACE "Info:本Jump节点的目的节点没有找到"
	End If
	
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
    			LastStateUuid = CurrProcessNode.Attributes.getNamedItem("ID").nodeValue
    		
			
		else
			Trace "是过渡状态"	
		end if	
		NextModuleName =  "state"
		Exit Sub
	
	Case "Operation":
		NextModuleName =  "Operation"
		Exit Sub
	Case "Jump":
		NextModuleName =  "Jump"
		Exit Sub		
	Case Else:
		Trace "Err!子节点类型:" & CurrProcessNode.baseName & "无法处理！"
		
		Exit Sub	
	End Select
End Sub

Sub OnFrontEndEvent(EvtName)
On Error Resume Next
	Trace "在Jump中捕捉到不合适FrontEnd事件：" & EvtName	
End Sub

Sub OnSoftPhoneEvent(EvtName)
On Error Resume Next
	Trace "在Jump中捕捉到不合适SoftPhone命令：" & EvtName	
End Sub


Sub OnTimeOut()
	Trace "Info:Jump收到不正确事件：TimeOut"
	
End Sub