'����״̬�ڵ�
Dim sName
Sub OnLoad()
	'��Trace����������һ��״̬��
	'Ȼ�����Ϸ��ذ�ť״̬��Ϣ
	Dim sXMLState
	sXMLState = "<STATUS>"
        sXMLState = sXMLState & "<Button Name=""Hook"" Title=""OffHook"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""Hold"" Title=""Hold"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""Transfer"" Title=""Transfer"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""DialOut"" Title=""DialOut"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""Consultation"" Title=""Consultation"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""Auto"" Title=""Auto"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""OutPhone"" Title=""OutPhone"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""Fax"" Title=""Fax"" Enable=""1""/>"
 	
 	sXMLState = sXMLState & "<Button Name=""Pause"" Title=""Pause"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""Conference"" Title=""Conference"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""Play"" Title=""Play"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""ReceiveDTMF"" Title=""ReceiveDTMF"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""Listen"" Title=""Listen"" Enable=""1""/>"
        sXMLState = sXMLState & "<Button Name=""RecordSeat"" Title=""RecordSeat"" Enable=""1""/>"
	sXMLState = sXMLState & "<Button Name=""Disconnect"" Title=""Disconnect"" Enable=""1""/>"
	sXMLState = sXMLState & "<Button Name=""RopCall"" Title=""RopCall"" Enable=""1""/>"
	sXMLState = sXMLState & "</STATUS>"


	Dim NodeX
	Dim xNodeTmp
	Dim NodeList
	Dim Name
	
	sName = CurrProcessNode.Attributes.getNamedItem("Name").nodeValue
	TRACE "Info:OnLoad()--- state.bas"
	
	'TRACE ֻ��TRACE��һ���״̬
	TRACE "Info:����״̬����>" & sName
	
	'�ж��Ƿ�Ϊ��ʱ״̬
	'��״̬�������Ƿ���"��"
	If InStr(1,sName,"��") = 0 Then
		'������ʱ״̬��
		'KILL Timer
		Trace "Info:Timer Disabled..."
		SetTimerEnabled False
	End If

	
	
	'ȡ����״̬����Ӧ�İ�ť״̬����
	'����ϳ�һ��XML�ı�
	'�������״̬XML�ı��Ϸ���SoftPhone�����ȥ
	
	Dim filter
	Dim filter1
	filter = "./Buttons/*"
	
	Dim sIID
	sIID = CurrProcessNode.Attributes.getNamedItem("ID").nodeValue
	
	Set NodeList = CurrProcessNode.selectNodes(filter)
	
	Trace m_StatePath
	
	'��ȡ��һ��״̬XML����
	
	Dim xDocTmp
	Set xDocTmp = CreateObject("MSXML.DOMDocument")
	'xDocTmp.Load "D:\AgentInterpretor\Bin\state.xml"
	
	'xDocTmp.Load m_StatePath
	xDocTmp.LoadXML sXMLState
	
	
    	Set xNodeTmp = xDocTmp.ChildNodes(0)
    	
    	Dim NodeNewAttr
    	Set NodeNewAttr = xDocTmp.createAttribute("Name")
    	
        NodeNewAttr.nodeValue = sName
        xNodeTmp.Attributes.setNamedItem NodeNewAttr

	Set NodeNewAttr = xDocTmp.createAttribute("ID")
    	
        NodeNewAttr.nodeValue = sIID
        xNodeTmp.Attributes.setNamedItem NodeNewAttr

	
	For each xNode in NodeList
		'���ڷ����в������Ӧ��Button
		
		Name = xNode.Attributes.getNamedItem("Name").nodeValue
		filter1 = "//STATUS/*[@Name=" & Chr(34) & Name & Chr(34) & "]"
		'Trace "Info:" & filter1
		Set xNodeTmp = xDocTmp.selectSingleNode(filter1)
		
		If Not xNodeTmp Is Nothing Then
			'���ҵ��󣬸�����Ӧ������ֵ
			
			xNodeTmp.Attributes.getNamedItem("Title").nodeValue = xNode.Attributes.getNamedItem("Title").nodeValue
			
			xNodeTmp.Attributes.getNamedItem("Enable").nodeValue = xNode.Attributes.getNamedItem("Enable").nodeValue
			'Trace "Info:Enable=" & xNodeTmp.Attributes.getNamedItem("Enable").nodeValue
		End if
	Next
	
	
	'Ȼ������µ�״̬�Ϸ���SoftPhone�����
	
	
	RaiseStatus xDocTmp.xml
	
	
	'�ж��Ƿ�Ϊ��ʱ״̬
	'��״̬�������Ƿ���"��"
	If InStr(1,sName,"��") <> 0 and InStr(1,sName,"����������״̬") = 0 and InStr(1,sName,"������״̬") = 0 Then
		'����ʱ״̬��
		'����Timer���ж��Ƿ�ʱ
		Trace "Info:Timer Enabled..."
		SetTimerInterval 60
		SetTimerEnabled True
	End If
	
	'�����ִֹͣ���κζ������ȴ�SoftPhone����FrontEnd��Ϣ�Ĵ���
	
End Sub

Sub OnFrontEndEvent()
	EvtName = CurrReceivedMsg
	On Error Resume Next
	On Error Resume Next
	Trace "Info:���յ�FrontEnd�¼�-->" & EvtName	
	
	'�����ӽڵ����Ƿ��и��¼�
	'����У���ʾ��Ҫ�Ը��¼�������Ӧ
	'���û�У���ʾ�ڱ�״̬�£����¼�Ϊ��Ч�¼�
	
	
	
	Dim NodeList 
	Dim xNode
	Dim Evt
	Dim Temp
	Dim bFound
	
	bFound = False
	'��ȡ����Ϣ������
	
	xDocFrontEnd.loadXML EvtName
	Set xNode = xDocFrontEnd.selectSingleNode("//MSG")
	Evt = xNode.attributes.getNamedItem("EVT").nodeValue
	
	'��ѭ���鿴�Ƿ������������Ϣƥ�����Ϣ����
	
	Dim filter
	
	for each xNode in CurrProcessNode.childNodes
		
		if  xNode.attributes.length = 0 then
		elseif  xNode.attributes.getNamedItem("EVT") is Nothing then
			
		elseif InStr(xNode.attributes.getNamedItem("EVT").nodeValue,Evt) > 0 then 
			bFound = True
			Set CurrProcessNode = xNode
			exit for
		end if
	next

	If bFound = True Then
		'Set CurrProcessNode = CurrProcessNode.childNodes.Item(0)
		'�����������ж�����һ����֧
		
		
		Set xNodeList = CurrProcessNode.childNodes	   
    		
    		bFound = false
    			
    		If xNodeList.Length > 1 Then
    			
			For each xNode in xNodeList
		
		    		if xNode.selectNodes("./@����").Length = 0 then
					ConditionName = "./@Condition"
				else
					ConditionName = "./@����"
				end if
		
		
			    	If xNode.selectNodes(ConditionName).Length <> 0 Then
			    		
					if xNode.selectNodes(ConditionName).Length <> 0 Then
						
						Condition = Trim(xNode.selectSingleNode(ConditionName).nodeValue)
					end if
					
					
				    	If  Condition= "" Then
				    		Condition = True
				    	End If
					Trace xNode.NodeName & "..." & Eval(Condition ) & "..." & Condition
				    	If Eval(Condition) = true  Then
				    		Set CurrProcessNode = xNode
				    		bFound = true
						exit for
				    	End If
				Else
						Set CurrProcessNode = xNode
				    		bFound = true
						exit For
						    	
				End If
				   
			Next
		ElseIf xNodeList.Length = 1 Then
			
			bFound = true
			Set CurrProcessNode = xNodeList.Item(0)
		ElseIf xNodeList.Length = 0 Then
			
			bFound = false
			Set CurrProcessNode = Nothing
		End If
		
		If(bFound = true) Then
			JumpToNode CurrProcessNode
		Else
			TRACE "Info:�������޺��ʵ��ӽڵ㣡"
		End If

		'JumpToNode CurrProcessNode
		
			
		Exit Sub
	Else
		TRACE "Info:���¼�û�ж�Ӧ�Ĵ�������"
		RaiseEvents "Not Avaliable Evt" & Evt ,EvtName	
	End If
		
End Sub

Sub OnSoftPhoneEvent()
	'On Error Resume Next
	EvtName = CurrReceivedMsg
	Trace "Info:���յ�SoftPhone�¼�-->" & EvtName	
	
	
	
	'�����ӽڵ����Ƿ��и��¼�
	'����У���ʾ��Ҫ�Ը��¼�������Ӧ
	'���û�У���ʾ�ڱ�״̬�£����¼�Ϊ��Ч�¼�
	
	Dim NodeList 
	Dim xNode
	Dim Cmd
	Dim Temp
	Dim bFound
	
	bFound = False
	'��ȡ����Ϣ������
	
	xDocSoftPhone.loadXML EvtName
	Set xNode = xDocSoftPhone.selectSingleNode("//MSG")
	Cmd = xNode.attributes.getNamedItem("CMD").nodeValue
	
	'�ж��Ƿ�ΪLogOff�¼�
	'�����LogOff�¼�����һ���Ȼ�����Main.bas,�����½���δ��¼״̬
	If Cmd = "LOGOFF" Then
		SendMsg2FrontEnd
		NextModuleName =  "Main"
		Exit Sub
	End If
	
	
	'��ѭ���鿴�Ƿ������������Ϣƥ�����Ϣ����
	
	Dim filter


	for each xNode in CurrProcessNode.childNodes
		
		if  xNode.attributes.length = 0 then
		elseif  xNode.attributes.getNamedItem("CMD") is Nothing then
			
		elseif InStr(xNode.attributes.getNamedItem("CMD").nodeValue,Cmd) > 0 then 
			bFound = True
			Set CurrProcessNode = xNode
			exit for
		end if
	next
		
	
	If bFound = True Then
		SetTimerEnabled False
		Trace "Info:Timer Disabled..."
		Set CurrProcessNode = CurrProcessNode.childNodes.Item(0)
		JumpToNode CurrProcessNode
		'��Event�Ϸ���SoftPhone�����
		RaiseEvents Cmd, EvtName	
		Exit Sub
	Else
		TRACE "Info:���¼�û�ж�Ӧ�Ĵ�������"
		RaiseEvents "Not Available Command:" & Cmd ,EvtName
	End If
End Sub

Sub JumpToNode(xNode)
	On Error Resume Next
	'���ݽ�����ͽ�����ת
	sName = CurrProcessNode.Attributes.getNamedItem("Name").nodeValue

	Select Case CurrProcessNode.baseName
	Case "state":
		'����LastStateUuid��
		
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

Sub OnTimeOut()
	Trace "Info:��״̬TimeOut���ص���һ״̬������" & sName
	
	SetTimerEnabled False
	'�ȶ�λ��һ״̬��Node,����LastStateUuid����
	
	Trace "Info:" & LastStateUuid
	
	Dim filter
	if (InStr(1,sName,"���н�����״̬") <>0)  or (InStr(1,sName,"���н�����״̬") <>0) then
		Dim xml
		xml = "<MSG CMD=""OFFHOOKFAIL""></MSG>"
		SendMsgToCtiServer xml	
	end if
	
	if (InStr(1,sName,"���������״̬") <>0) or (InStr(1,sName,"������״̬") <>0) then
		
		xml = "<MSG CMD=""OFFHOOKFAIL""></MSG>"
		SendMsgToCtiServer xml
		
		'LastStateUuid = "F3FEEB8C80E64B0D92AFC4EC84E6C4F6" 'ת�����
			
	end if

	if (InStr(1,sName,"ת����״̬") <>0) then
		
		xml = "<MSG CMD=""OFFHOOKFAIL""></MSG>"
		SendMsgToCtiServer xml
		
		'LastStateUuid = "F3FEEB8C80E64B0D92AFC4EC84E6C4F6" 'ת�����
			
	end if

	if (InStr(1,sName,"�Ⲧ��״̬") <>0) or (InStr(1,sName,"�����Ⲧ��״̬") <>0) then
		
		xml = "<MSG CMD=""OFFHOOKFAIL""></MSG>"
		SendMsgToCtiServer xml		
		'LastStateUuid = "F3FEEB8C80E64B0D92AFC4EC84E6C4F6" 'ת�����
			
	end if

	if (InStr(1,sName,"����������״̬") <>0) then 
		xml = "<MSG CMD=""OFFHOOKFAIL""></MSG>"
		SendMsgToCtiServer xml	
	end if

	if (InStr(1,sName,"���ʹ�����״̬") <>0)  then
		xml = "<MSG CMD=""ONHOOK""></MSG>"
		SendMsgToCtiServer xml
	end if

	
	
	filter = "//*[@ID=""" & LastStateUuid & """]"
	
	Set CurrProcessNode = xDoc.selectSingleNode(filter)
	
	JumpToNode CurrProcessNode
	
	'��Event�Ϸ���SoftPhone�����
	RaiseEvents sName , sName & "TimeOut"	
	
	
End Sub