'�����ڵ�
Sub OnLoad
	'�Ȳ��ұ��ڵ��Ӧ�Ĳ����ڵ�
	Trace "Info:OnLoad()--- Operation.bas"
	Dim filter
	Dim OperationNode, tmpNode
	filter = "./Operations"
	
	Set OperationNode = CurrProcessNode.SelectSingleNode(filter) 'ȡ�ò����ڵ�
	
	On Error Resume Next
	Dim i '��������
	
	If OperationNode Is Nothing Then '�ж��Ƿ���ڲ����ڵ�
		
	Else
	    Trace "Info:Yes Operation;Number is " & OperationNode.ChildNodes.Length
	    For i = 1 To OperationNode.ChildNodes.Length '�в����ڵ㣬 ��Բ����ڵ����ѭ������        
	    	
	    	Trace "Info:ִ�в���:" & OperationNode.ChildNodes(i-1).Attributes.getNamedItem("Expression").nodeValue
	    	Execute OperationNode.ChildNodes(i-1).Attributes.getNamedItem("Expression").nodeValue '���в���
										    		
	    	
	    Next
	End If
	
	'ִ�������ת����һ���ڵ�
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
	'���ݽ�����ͽ�����ת
	Dim sName
	sName = CurrProcessNode.Attributes.getNamedItem("Name").nodeValue

	Select Case CurrProcessNode.baseName
	Case "state":
		If InStr(1,sName,"��") = 0 Then
			Trace "���ǹ���״̬"
			
		else
			Trace "�ǹ���״̬"
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
		Trace "Info:Err!�ӽڵ�����:" & CurrProcessNode.baseName & "�޷�����"
		
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