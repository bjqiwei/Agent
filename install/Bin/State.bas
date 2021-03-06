Attribute VB_Name = "Module1"
'处理状态节点
Dim sName
Sub OnLoad()
        '先Trace出正处在哪一种状态；
        '然后向上返回按钮状态消息
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
        Trace "Info:OnLoad()--- state.bas"
        
        'TRACE 只是TRACE出一般的状态
        Trace "Info:进入状态－－>" & sName
        
        '判断是否为临时状态
        '即状态命名中是否含有"中"
        If InStr(1, sName, "中") = 0 Then
                '不是临时状态；
                'KILL Timer
                Trace "Info:Timer Disabled..."
                SetTimerEnabled False
        End If

        
        
        '取出本状态所对应的按钮状态集合
        '并组合成一个XML文本
        '最后把这个状态XML文本上返到SoftPhone外壳中去
        
        Dim filter
        Dim filter1
        filter = "./Buttons/*"
        
        Dim sIID
        sIID = CurrProcessNode.Attributes.getNamedItem("ID").nodeValue
        
        Set NodeList = CurrProcessNode.selectNodes(filter)
        
        Trace m_StatePath
        
        '先取出一个状态XML范本
        
        Dim xDocTmp
        Set xDocTmp = CreateObject("MSXML.DOMDocument")
        'xDocTmp.Load "D:\AgentInterpretor\Bin\state.xml"
        
        'xDocTmp.Load m_StatePath
        xDocTmp.loadXML sXMLState
        
        
        Set xNodeTmp = xDocTmp.childNodes(0)
        
        Dim NodeNewAttr
        Set NodeNewAttr = xDocTmp.createAttribute("Name")
        
        NodeNewAttr.nodeValue = sName
        xNodeTmp.Attributes.setNamedItem NodeNewAttr

        Set NodeNewAttr = xDocTmp.createAttribute("ID")
        
        NodeNewAttr.nodeValue = sIID
        xNodeTmp.Attributes.setNamedItem NodeNewAttr

        
        For Each xNode In NodeList
                '先在范本中查找相对应的Button
                
                Name = xNode.Attributes.getNamedItem("Name").nodeValue
                filter1 = "//STATUS/*[@Name=" & Chr(34) & Name & Chr(34) & "]"
                'Trace "Info:" & filter1
                Set xNodeTmp = xDocTmp.selectSingleNode(filter1)
                
                If Not xNodeTmp Is Nothing Then
                        '查找到后，更新相应的属性值
                        
                        xNodeTmp.Attributes.getNamedItem("Title").nodeValue = xNode.Attributes.getNamedItem("Title").nodeValue
                        
                        xNodeTmp.Attributes.getNamedItem("Enable").nodeValue = xNode.Attributes.getNamedItem("Enable").nodeValue
                        'Trace "Info:Enable=" & xNodeTmp.Attributes.getNamedItem("Enable").nodeValue
                End If
        Next
        
        
        '然后把最新的状态上返到SoftPhone外壳中
        
        
        RaiseStatus xDocTmp.xml
        
        
        '判断是否为临时状态
        '即状态命名中是否含有"中"
        If InStr(1, sName, "中") <> 0 And InStr(1, sName, "磋商外线中状态") = 0 And InStr(1, sName, "磋商中状态") = 0 Then
                '是临时状态；
                '启动Timer，判断是否超时
                Trace "Info:Timer Enabled..."
                SetTimerInterval 30
                SetTimerEnabled True
        End If
        
        '下面会停止执行任何动作，等待SoftPhone或者FrontEnd消息的触发
        
End Sub

Sub OnFrontEndEvent()
        EvtName = CurrReceivedMsg
        On Error Resume Next
        On Error Resume Next
        Trace "Info:接收到FrontEnd事件-->" & EvtName
        
        '查找子节点中是否有该事件
        '如果有，表示需要对该事件作出响应
        '如果没有，表示在本状态下，该事件为无效事件
        
        
        
        Dim NodeList
        Dim xNode
        Dim Evt
        Dim Temp
        Dim bFound
        
        bFound = False
        '先取得消息的类型
        
        xDocFrontEnd.loadXML EvtName
        Set xNode = xDocFrontEnd.selectSingleNode("//MSG")
        Evt = """" & xNode.Attributes.getNamedItem("EVT").nodeValue & """"
        
        '再循环查看是否有类型与此消息匹配的消息定义
        
        Dim filter
        
        For Each xNode In CurrProcessNode.childNodes
                
                If xNode.Attributes.length = 0 Then
                ElseIf xNode.Attributes.getNamedItem("EVT") Is Nothing Then
                        
                'ElseIf InStr(xNode.Attributes.getNamedItem("EVT").nodeValue, Evt) > 0 Then
                Else
                    Trace "Info1..." & xNode.Attributes.getNamedItem("EVT").nodeValue
                    Trace "Info2..." & Evt
                    If xNode.Attributes.getNamedItem("EVT").nodeValue = Evt Then
                            bFound = True
                            Set CurrProcessNode = xNode
                            Exit For
                    End If
                End If
        Next

        If bFound = True Then
                'Set CurrProcessNode = CurrProcessNode.childNodes.Item(0)
                '根据条件来判断走哪一个分支
                
                
                Set xNodeList = CurrProcessNode.childNodes
                
                bFound = False
                        
                If xNodeList.length > 1 Then
                        
                        For Each xNode In xNodeList
                
                                If xNode.selectNodes("./@条件").length = 0 Then
                                        ConditionName = "./@Condition"
                                Else
                                        ConditionName = "./@条件"
                                End If
                
                
                                If xNode.selectNodes(ConditionName).length <> 0 Then
                                        
                                        If xNode.selectNodes(ConditionName).length <> 0 Then
                                                
                                                Condition = Trim(xNode.selectSingleNode(ConditionName).nodeValue)
                                        End If
                                        
                                        
                                        If Condition = "" Then
                                                Condition = True
                                        End If
                                        Trace xNode.NodeName & "..." & Eval(Condition) & "..." & Condition
                                        If Eval(Condition) = True Then
                                                Set CurrProcessNode = xNode
                                                bFound = True
                                                Exit For
                                        End If
                                Else
                                                Set CurrProcessNode = xNode
                                                bFound = True
                                                Exit For
                                                        
                                End If
                                   
                        Next
                ElseIf xNodeList.length = 1 Then
                        
                        bFound = True
                        Set CurrProcessNode = xNodeList.Item(0)
                ElseIf xNodeList.length = 0 Then
                        
                        bFound = False
                        Set CurrProcessNode = Nothing
                End If
                
                If (bFound = True) Then
                        JumpToNode CurrProcessNode
                Else
                        Trace "Info:！！！无合适的子节点！"
                End If

                'JumpToNode CurrProcessNode
                
                        
                Exit Sub
        Else
                Trace "Info:本事件没有对应的处理流程"
                RaiseEvents "Not Avaliable Evt" & Evt, EvtName
        End If
                
End Sub

Sub OnSoftPhoneEvent()
        'On Error Resume Next
        EvtName = CurrReceivedMsg
        Trace "Info:接收到SoftPhone事件-->" & EvtName
        
        
        
        '查找子节点中是否有该事件
        '如果有，表示需要对该事件作出响应
        '如果没有，表示在本状态下，该事件为无效事件
        
        Dim NodeList
        Dim xNode
        Dim Cmd
        Dim Temp
        Dim bFound
        
        bFound = False
        '先取得消息的类型
        
        xDocSoftPhone.loadXML EvtName
        Set xNode = xDocSoftPhone.selectSingleNode("//MSG")
        Cmd = xNode.Attributes.getNamedItem("CMD").nodeValue
        
        '判断是否为LogOff事件
        '如果是LogOff事件，则挂机，然后进入Main.bas,即重新进入未登录状态
        If Cmd = "LOGOFF" Then
                SendMsg2FrontEnd
                NextModuleName = "Main"
                Exit Sub
        End If
        
        
        '再循环查看是否有类型与此消息匹配的消息定义
        
        Dim filter


        For Each xNode In CurrProcessNode.childNodes
                
                If xNode.Attributes.length = 0 Then
                ElseIf xNode.Attributes.getNamedItem("CMD") Is Nothing Then
                        
                ElseIf InStr(xNode.Attributes.getNamedItem("CMD").nodeValue, Cmd) > 0 Then
                        bFound = True
                        Set CurrProcessNode = xNode
                        Exit For
                End If
        Next
                
        
        If bFound = True Then
                SetTimerEnabled False
                Trace "Info:Timer Disabled..."
                Set CurrProcessNode = CurrProcessNode.childNodes.Item(0)
                JumpToNode CurrProcessNode
                '把Event上返到SoftPhone外壳中
                RaiseEvents Cmd, EvtName
                Exit Sub
        Else
                Trace "Info:本事件没有对应的处理流程"
                RaiseEvents "Not Available Command:" & Cmd, EvtName
        End If
End Sub

Sub JumpToNode(xNode)
        On Error Resume Next
        '根据结点类型进行跳转
        sName = CurrProcessNode.Attributes.getNamedItem("Name").nodeValue

        Select Case CurrProcessNode.baseName
        Case "state":
                '设置LastStateUuid：
                
                If InStr(1, sName, "中") = 0 Then
                        Trace "不是过渡状态"
                        LastStateUuid = CurrProcessNode.Attributes.getNamedItem("ID").nodeValue
                Else
                        Trace "是过渡状态"
                End If
                NextModuleName = "state"
                Exit Sub
        Case "Operation":
                NextModuleName = "Operation"
                Exit Sub
        Case "Jump":
                NextModuleName = "Jump"
                Exit Sub
        Case Else:
                Trace "Err!子节点类型:" & CurrProcessNode.baseName & "无法处理！"
                
                Exit Sub
        End Select
End Sub

Sub OnTimeOut()
        Trace "Info:本状态TimeOut：回到上一状态。。。" & sName
        
        SetTimerEnabled False
        '先定位上一状态的Node,根据LastStateUuid来定
        
        Trace "Info:" & LastStateUuid
        
        Dim filter
        If (InStr(1, sName, "呼叫进入中状态") <> 0) Or (InStr(1, sName, "呼叫进入中状态") <> 0) Then
                Dim xml
                xml = "<MSG CMD=""OFFHOOKFAIL""></MSG>"
                SendMsgToCtiServer xml
        End If
        
        If (InStr(1, sName, "请求监听中状态") <> 0) Or (InStr(1, sName, "磋商中状态") <> 0) Then
                
                xml = "<MSG CMD=""OFFHOOKFAIL""></MSG>"
                SendMsgToCtiServer xml
                
                'LastStateUuid = "F3FEEB8C80E64B0D92AFC4EC84E6C4F6" '转入空闲
                        
        End If

        If (InStr(1, sName, "转接中状态") <> 0) Then
                
                xml = "<MSG CMD=""OFFHOOKFAIL""></MSG>"
                SendMsgToCtiServer xml
                
                'LastStateUuid = "F3FEEB8C80E64B0D92AFC4EC84E6C4F6" '转入空闲
                        
        End If

        If (InStr(1, sName, "外拨中状态") <> 0) Or (InStr(1, sName, "会议外拨中状态") <> 0) Then
                
                xml = "<MSG CMD=""OFFHOOKFAIL""></MSG>"
                SendMsgToCtiServer xml
                'LastStateUuid = "F3FEEB8C80E64B0D92AFC4EC84E6C4F6" '转入空闲
                        
        End If

        If (InStr(1, sName, "磋商外线中状态") <> 0) Then
                xml = "<MSG CMD=""OFFHOOKFAIL""></MSG>"
                SendMsgToCtiServer xml
        End If

        If (InStr(1, sName, "发送传真中状态") <> 0) Then
                xml = "<MSG CMD=""ONHOOK""></MSG>"
                SendMsgToCtiServer xml
        End If

        
        
        filter = "//*[@ID=""" & LastStateUuid & """]"
        
        Set CurrProcessNode = xDoc.selectSingleNode(filter)
        
        JumpToNode CurrProcessNode
        
        '把Event上返到SoftPhone外壳中
        RaiseEvents sName, sName & "TimeOut"
        
        
End Sub
