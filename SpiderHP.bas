Attribute VB_Name = "SpiderHP"
Option Explicit
'������
Sub Main()
    Dim strText As String
    With CreateObject("MSXML2.XMLHTTP") 'CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "POST", "", False
        .setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        .setRequestHeader "Referer", ""
        .Send
        strText = .responsetext
        Debug.Print strText
    End With
End Sub


'Javascript���ʽ��ֵ
Function JSEval(strText As String) As String
    With CreateObject("MSScriptControl.ScriptControl")
        .Language = "javascript"
        JSEval = .Eval(strText)
    End With
End Function
'urlת��
Function encodeURI(strText As String) As String
    With CreateObject("msscriptcontrol.scriptcontrol")
        .Language = "JavaScript"
        encodeURI = .Eval("encodeURIComponent('" & strText & "');")
    End With
End Function
'������ת��ָ��������ı�
Function ByteToStr(arrByte, strCharset As String) As String
    With CreateObject("Adodb.Stream")
        .Type = 1 'adTypeBinary
        .Open
        .write arrByte
        .Position = 0
        .Type = 2 'adTypeText
        .Charset = strCharset
        ByteToStr = .Readtext
        .Close
    End With
End Function
'�ı���ָ������תΪ������
Function StrToByte(strText As String, strCharset As String, Optional Bom As Boolean = False)
    With CreateObject("adodb.stream")
        .Type = 2 'adTypeText
        .Charset = strCharset
        .Open
        .Writetext strText
        .Position = 0
        .Type = 1 'adTypeBinary
        If Not Bom Then
            If LCase(strCharset) = "unicode" Then
                .Position = 2
            ElseIf LCase(strCharset) = "utf-8" Then
                .Position = 3
            End If
        End If
        StrToByte = .Read
    End With
End Function
'��������ת���ļ�
Sub ByteToFile(arrByte, strFileName As String)
    With CreateObject("Adodb.Stream")
        .Type = 1 'adTypeBinary
        .Open
        .write arrByte
        .SaveToFile strFileName, 2 'adSaveCreateOverWrite
        .Close
    End With
End Sub
'�ı�������������
Sub CopyToClipbox(strText As String)
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText strText
        .PutInClipboard
    End With
End Sub
'ִ��js����ǡ���2
Function EvalByHtml(strText As String) As String
    With CreateObject("htmlfile")
        .write "<html><script></script></html>"
        EvalByHtml = CallByName(.parentwindow, "eval", VbMethod, strText)
    End With
End Function
'url����
Function encodeURIByHtml(strText As String) As String
    With CreateObject("htmlfile")
        .write "<html><script></script></html>"
        encodeURIByHtml = CallByName(.parentwindow, "encodeURIComponent", VbMethod, strText)
    End With
End Function
'unicode ����
Function unescape(strTobecoded As String) As String
    With CreateObject("msscriptcontrol.scriptcontrol")
        .Language = "JavaScript"
        unescape = .Eval("unescape('" & strTobecoded & "');")
    End With
End Function
'����
Function Reg()
    With CreateObject("VBScript.Regexp")
    .Global = True
    .Pattern = "{""b"":\d+,""g"":\d+,""n"":""([^""]*)"",""u"":(\d+)}"
    For Each RegMatch In .Execute(strText)
        n = n + 1
        arrData(n, 1) = RegMatch.submatches(0)
        arrData(n, 2) = RegMatch.submatches(1)
    Next
End With
End Function


Function Base64ToByte(strBase As String) As Byte()
    With CreateObject("Microsoft.XMLDOM")
        With .createElement("a")
            .DataType = "bin.base64"
            .Text = strBase
            Base64ToByte = .nodeTypedValue
        End With
    End With
End Function

'.Option(6) = False ' ��ֹ�ض����Ի�ȡԭ��ҳ��Ϣ
'.SetProxy 2, "218.75.100.114:8080" '����
