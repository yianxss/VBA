Attribute VB_Name = "SpiderHP"
Option Explicit
'主函数
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


'Javascript表达式求值
Function JSEval(strText As String) As String
    With CreateObject("MSScriptControl.ScriptControl")
        .Language = "javascript"
        JSEval = .Eval(strText)
    End With
End Function
'url转码
Function encodeURI(strText As String) As String
    With CreateObject("msscriptcontrol.scriptcontrol")
        .Language = "JavaScript"
        encodeURI = .Eval("encodeURIComponent('" & strText & "');")
    End With
End Function
'流数据转成指定编码的文本
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
'文本按指定编码转为流数据
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
'二进制流转成文件
Sub ByteToFile(arrByte, strFileName As String)
    With CreateObject("Adodb.Stream")
        .Type = 1 'adTypeBinary
        .Open
        .write arrByte
        .SaveToFile strFileName, 2 'adSaveCreateOverWrite
        .Close
    End With
End Sub
'文本拷贝到剪贴板
Sub CopyToClipbox(strText As String)
    With CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        .SetText strText
        .PutInClipboard
    End With
End Sub
'执行js表达是——2
Function EvalByHtml(strText As String) As String
    With CreateObject("htmlfile")
        .write "<html><script></script></html>"
        EvalByHtml = CallByName(.parentwindow, "eval", VbMethod, strText)
    End With
End Function
'url编码
Function encodeURIByHtml(strText As String) As String
    With CreateObject("htmlfile")
        .write "<html><script></script></html>"
        encodeURIByHtml = CallByName(.parentwindow, "encodeURIComponent", VbMethod, strText)
    End With
End Function
'unicode 解码
Function unescape(strTobecoded As String) As String
    With CreateObject("msscriptcontrol.scriptcontrol")
        .Language = "JavaScript"
        unescape = .Eval("unescape('" & strTobecoded & "');")
    End With
End Function
'正则
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

'.Option(6) = False ' 禁止重定向，以获取原网页信息
'.SetProxy 2, "218.75.100.114:8080" '代理
