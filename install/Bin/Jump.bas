'Jump�ڵ�

Sub OnLoad()
	'�ҵ�Destination�ڵ�
	
	 Trace "Info:OnLoad()--- Jump.bas"
	Dim NodeList
	Dim destination
	Dim xNodeTmp
	Dim filter
	Dim bFound 
	
	bFound = False
	
	'��ȡ��Destination����
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
		TRACE "Info:��Jump�ڵ��Ŀ�Ľڵ�û���ҵ�"
	End If
	
End Sub

Sub JumpToNode(xNode)
	On Error Resume Next
	'���ݽ�����ͽ�����ת
	Dim sName
	sName = CurrProcessNode.Attributes.getNamedItem("Name").nodeValue

	Select Case CurrProcessNode.baseName
	Case "state":
		If InStr(1,sName,"��") = 0 Then
                        Trace "���ǹ���״̬"
    			LastStateUuid = CurrProcessNode.Attributes.getNamedItem("ID").nodeValue
    		
			
		else
			Trace "�ǹ���״̬"	
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
		Trace "Err!�ӽڵ�����:" & CurrProcessNode.baseName & "�޷�����"
		
		Exit Sub	
	End Select
End Sub

Sub OnFrontEndEvent(EvtName)
On Error Resume Next
	Trace "��Jump�в�׽��������FrontEnd�¼���" & EvtName	
End Sub

Sub OnSoftPhoneEvent(EvtName)
On Error Resume Next
	Trace "��Jump�в�׽��������SoftPhone���" & EvtName	
End Sub


Sub OnTimeOut()
	Trace "Info:Jump�յ�����ȷ�¼���TimeOut"
	
End Sub