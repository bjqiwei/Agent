'���̵Ľ����
'���ȶ�λ���̵ĵ�һ���ڵ㣬 Ȼ��ȷ���ڵ�����ʣ� ������Ӧ�Ľڵ㴦��ű�
Sub OnLoad()
    Trace "Info:OnLoad()--- Main.bas"
    '���ȶ�λxDoc�еĵ�һ���ڵ�
    
  
    Dim NodesList
   
    Dim filter
    
    '����ȡ�õ�һ��״̬�ڵ㣻��Ȼ��AgentRun�µĵ�һ��state�ڵ�
    filter = "//AgentRun/state"
    Set NodesList = xDoc.selectNodes(filter)
    
    '��CurrProcessNode��λΪ��һ��state�ڵ㣻
    Set CurrProcessNode = NodesList.Item(0)
    
    
    
    '��ת����һ���ڵ�
    JumpToNode CurrProcessNode
    
    
    
End Sub



Sub JumpToNode(xNode)
	On Error Resume Next
	'���ݽ�����ͽ�����ת
	Select Case CurrProcessNode.baseName
	Case "state":
		'����LastStateUuid��
    		LastStateUuid = CurrProcessNode.Attributes.getNamedItem("ID").nodeValue
    		
		NextModuleName =  "state"
		Exit Sub
	
	Case "Operation":
		NextModuleName =  "Operation"
		Exit Sub		
	Case Else:
		Trace "Err!�ӽڵ�����:" & CurrProcessNode.baseName & "�޷�����"
		
		Exit Sub	
	End Select
End Sub

Sub OnFrontEndEvent(EvtName)
On Error Resume Next
	Trace "��Main�в�׽��������FrontEnd�¼���" & EvtName	
End Sub

Sub OnSoftPhoneEvent(EvtName)
On Error Resume Next
	Trace "��Main�в�׽��������SoftPhone���" & EvtName	
End Sub


Sub OnTimeOut()
	Trace "Info:�յ�����ȷ�¼���TimeOut"
	
End Sub