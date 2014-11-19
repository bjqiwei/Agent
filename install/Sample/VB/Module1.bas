Attribute VB_Name = "Module1"
Public DialOutMsg As String
Public LoginMsg As String
Public LoginPMsg As String
Public FaxOutMsg As String
Public FaxOutFileMsg As String
Global AgentID As String



Function ParseXML(Source As String, xmlDoc As MSXML.DOMDocument) As Boolean
  
  Dim errtext As String
  
    ParseXML = True

    Set xmlDoc = New DOMDocument
    
    With xmlDoc
'        .async = False
        .loadXML Source
    End With

    'xmlDoc.loadXML Source
    If xmlDoc.parseError.errorCode = 0 Then
        Exit Function
    End If
    With xmlDoc.parseError
        
        errtext = "document Parse Error:" & vbCrLf & _
        "Code: " & .errorCode & vbCrLf & _
        "Line: " & .Line & vbCrLf & _
        "lPos: " & .linepos & vbCrLf & _
        "Reason: " & .reason & vbCrLf & _
        "Src: " & .srcText & vbCrLf & _
        "fPos: " & .filepos
    End With
    ParseXML = False
    
    Set xmlDoc = Nothing
    
    
End Function
Public Function GetXMLValue(ByVal strName As String, ByVal StrSource As String, ByVal StrXMLCount As Integer, ByVal strCol As String) As String
    Dim m_XML As New MSXML.DOMDocument
    Dim strIndex As Integer
    On Error GoTo errHandle
    If StrXMLIndex > StrXMLCount Then
        GetXMLValue = ""
        Exit Function
    Else
        If ParseXML(StrSource, m_XML) = True Then
            For i = 1 To StrXMLCount
                If m_XML.childNodes.Item(0).childNodes.Item(i - 1).Attributes.getNamedItem("Name").nodeValue = strName Then
                    strIndex = i - 1
                    GetXMLValue = m_XML.childNodes.Item(0).childNodes.Item(strIndex).Attributes.getNamedItem(strCol).nodeValue
                    'Debug.Print GetXMLValue
                    Exit Function
                End If
            Next i
            
            GetXMLValue = ""
        End If
    End If
    
errHandle:
    GetXMLValue = "解析xml文件错误！！"
End Function

Public Function SetButtons(ByVal rstatus As String)
    EnableButton "Hook", frmMain.Hook, rstatus
    EnableButton "Hold", frmMain.Hold, rstatus
    EnableButton "Transfer", frmMain.Transfer, rstatus
    EnableButton "DialOut", frmMain.DialOut, rstatus
    EnableButton "Consultation", frmMain.Consultation, rstatus
    EnableButton "Auto", frmMain.Auto, rstatus
    EnableButton "OutPhone", frmMain.OutPhone, rstatus
    EnableButton "Fax", frmMain.Fax, rstatus
    EnableButton "Pause", frmMain.Pause, rstatus
    EnableButton "Conference", frmMain.Conference, rstatus
    EnableButton "Play", frmMain.Play, rstatus
    EnableButton "Listen", frmMain.Listen, rstatus
    EnableButton "Disconnect", frmMain.Disconnect, rstatus
    EnableButton "RopCall", frmMain.RopCall, rstatus
    
    
    
    SetButtonCaption "Hook", frmMain.Hook, rstatus
    SetButtonCaption "Hold", frmMain.Hold, rstatus
    SetButtonCaption "Transfer", frmMain.Transfer, rstatus
    SetButtonCaption "DialOut", frmMain.DialOut, rstatus
    SetButtonCaption "Consultation", frmMain.Consultation, rstatus
    SetButtonCaption "Auto", frmMain.Auto, rstatus
    SetButtonCaption "OutPhone", frmMain.OutPhone, rstatus
    SetButtonCaption "Fax", frmMain.Fax, rstatus
    SetButtonCaption "Pause", frmMain.Pause, rstatus
    SetButtonCaption "Conference", frmMain.Conference, rstatus
    SetButtonCaption "Play", frmMain.Play, rstatus
    SetButtonCaption "Listen", frmMain.Listen, rstatus
    SetButtonCaption "Disconnect", frmMain.Disconnect, rstatus
    SetButtonCaption "RopCall", frmMain.RopCall, rstatus
    
End Function

Public Sub EnableButton(ByVal strName As String, strButton As CommandButton, ByVal rstatus As String)
    If GetXMLValue(strName, rstatus, 16, "Enable") = "1" Then
        strButton.Enabled = True
    Else
        strButton.Enabled = False
    End If
End Sub

Public Sub SetButtonCaption(ByVal strName As String, strButton As CommandButton, ByVal rstatus As String)
    If Not GetXMLValue(strName, rstatus, 16, "Title") = "解析xml文件错误！！" Then
        strButton.Caption = GetXMLValue(strName, rstatus, 16, "Title")
    Else
        strButton.Caption = GetXMLValue(strName, rstatus, 16, "Title")
    End If
End Sub

Public Sub SetButtonOption(ByVal strTitle As String, ByVal strEnable As Integer, strButton As CommandButton)
    strButton.Caption = strTitle
    strButton.Enabled = strEnable
End Sub

Public Sub SetButtonInit()
    frmMain.Hook.Enabled = False
    frmMain.Hold.Enabled = False
    frmMain.Transfer.Enabled = False
    frmMain.DialOut.Enabled = False
    frmMain.Consultation.Enabled = False
    frmMain.Auto.Enabled = False
    frmMain.OutPhone.Enabled = False
    frmMain.Fax.Enabled = False
    frmMain.Pause.Enabled = False
    frmMain.Conference.Enabled = False
    frmMain.Play.Enabled = False
    frmMain.Listen.Enabled = False
    frmMain.Disconnect.Enabled = False
    frmMain.RopCall.Enabled = False
End Sub
