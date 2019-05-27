Attribute VB_Name = "ADOHP"
'数据是否包含标题行
Public Enum dbHDR
    Yes = 0
    No = 1
End Enum

'数据连接方式
Public Enum cnIMEX
    导入 = 0
    导出 = 1
    读写 = 2
End Enum

'Excel版本及类型
Public Enum cnType
    xls = 0
    xlsx = 1
    xlsm = 2
    Csv = 3
    accdb = 4
    mdb = 5
    txt = 6
    Auto = 7
End Enum

'查询类型
Public Enum QueryType
    Recordset = 0
    NonRes = 1
End Enum

'路径类型
Public Enum PathType
    CurrentPath = 0
    OtherPath = 1
End Enum

'返回数组是否包含标题
Public Enum ResultType
    onlyTitle = 0
    onlyBody = 1
    All = 2
End Enum



'连接数据库
Function ADOCNN(Optional dbName As String = "", _
                Optional ProType As cnType = cnType.Auto, _
                Optional bHDR As dbHDR = 0, _
                Optional cnIMEX As cnIMEX = 1, _
                Optional dbPwd As String = "" _
                )
                
    Dim cnn As Object
    Set cnn = CreateObject("Adodb.Connection")  '数据库连接
    
    '------------自动判断连接的类型------------
    Dim tempCnType As String
    If ProType = cnType.Auto Then ProType = getProviderByExtension(dbName)
    
    '------------是否包含标题行：bHDR----------
    Dim sHDR As String
    Select Case bHDR
        Case dbHDR.No
            sHDR = "No"
        Case dbHDR.Yes
            sHDR = "Yes"
    End Select
    
    
    '-----------处理自连接---------------------
    If Len(dbName) = 0 Then
        dbName = ThisWorkbook.FullName
    End If
    
    If InStr(dbName, "\") = 0 Then
        dbName = ThisWorkbook.Path & "\" & dbName
    End If
    
    
    '------------提供者：Provider--------------
    Dim strCnn As String
    Const ACE_PRO As String = "Provider = Microsoft.ACE.OlEDB.12.0;"
    Const JET_PRO As String = "Provider = Microsoft.JET.OlEDB.4.0;"
    
    Select Case ProType
        Case cnType.xlsx
            strCnn = ACE_PRO _
                    & "Extended Properties = 'Excel 12.0 Xml;HDR=" & sHDR _
                    & ";IMEX=" & cnIMEX & "';Data Source = " & dbName
        Case cnType.xlsm
            strCnn = ACE_PRO _
                    & "Extended Properties = 'Excel 12.0 Macro;HDR=" & sHDR _
                    & ";IMEX=" & cnIMEX & "';Data Source = " & dbName
        Case cnType.accdb
            strCnn = ACE_PRO & "Data Source = " & dbName _
                    & IIf(dbPwd = "", "", ";Jet OLEDB:Database Password=" & dbPwd)
        Case cnType.Csv, cnType.txt
            strCnn = ACE_PRO _
                    & "Extended Properties = 'Text;HDR=" & sHDR _
                    & ";FMT=Delimited';Data Source = " & dbName
        Case cnType.xls
            strCnn = JET_PRO _
                    & "Extended Properties = 'Excel 8.0;HDR=" & sHDR _
                    & ";IMEX=" & cnIMEX & "';Data Source = " & dbName
        Case cnType.mdb
            strCnn = JET_PRO & "Data Source = " & dbName _
                    & IIf(dbPwd = "", "", ";Jet OLEDB:Database Password=" & dbPwd)
        Case Else
            Debug.Print "类型不匹配,请检查"
    End Select
    
    On Error Resume Next
    cnn.Open strCnn
    If Err.Number <> 0 Then
        Debug.Print "数据库连接失败": ADOCNN = Nothing: Exit Function
    Else
        Set ADOCNN = cnn
    End If
End Function

'获取查询结果
Function SqlQuery(ByRef cn As Variant, sSql As String, _
                  Optional qTyper As QueryType = QueryType.Recordset)
    Dim rs As Object
    If cn Is Nothing Then Debug.Print "数据未成功连接": Exit Function
    Select Case qTyper
        Case QueryType.Recordset
            Set rs = CreateObject("ADODb.RecordSet")
            On Error Resume Next
            rs.Open sSql, cn, 3, 2
            If Err.Number <> 0 Then
                Debug.Print "查询失败,请检查SQL语法或者连接是否异常"
                Debug.Print Err.Description
                Set cn = Nothing: cn.Close
                Set rs = Nothing: Exit Function
            Else
                Set SqlQuery = rs
            End If
        
        Case QueryType.NonRes
            On Error Resume Next
            cn.Execute sSql
            If Err <> 0 Then
                SqlQuery = False: Debug.Print "查询失败!原因:" & Err.Description
            Else
                SqlQuery = True
            End If
             cn.Close: Set cn = Nothing
    End Select
End Function

'查询结果放入数组
Function RsToArr(ByRef rs As Variant, _
                 Optional resType As ResultType = ResultType.All)
    Dim i As Integer, j As Integer
    Dim brr(), arr()

    If rs.RecordCount = 0 Then
        RsToArr = 0: Exit Function
    Else
        arr = rs.GetRows
    End If
    Select Case resType
        Case ResultType.All
            ReDim brr(UBound(arr, 2) + 1, UBound(arr))
            For i = 0 To UBound(arr)
                For j = 0 To UBound(arr, 2)
                    brr(j + 1, i) = arr(i, j)
                Next
            Next
            For i = 0 To rs.Fields.Count - 1
                brr(0, i) = rs.Fields(i).Name
            Next
            RsToArr = brr
        Case ResultType.onlyBody
            ReDim brr(UBound(arr, 2), UBound(arr))
            For i = 0 To UBound(arr)
                For j = 0 To UBound(arr, 2)
                    brr(j, i) = arr(i, j)
                Next
            Next
            RsToArr = brr
        Case ResultType.onlyTitle
            ReDim brr(0, UBound(arr))
            For i = 0 To rs.Fields.Count - 1
                brr(0, i) = rs.Fields(i).Name
            Next
            RsToArr = brr
    End Select
End Function

'创建数据库文件
Function CreateAccDB(dbName As String, Optional sPath As PathType = 0)
    Dim sFilename As String, cnnString As String
    Dim cat As Object, sExtension As String
    Set cat = CreateObject("Adox.Catalog")
    
    cnnStringAr = selProvider(dbName)
    If sPath = PathType.CurrentPath Then
        sFilename = ThisWorkbook.Path & "\" & dbName & cnnStringAr(1)
    Else
        sFilename = dbName & cnnStringAr(1)
    End If
    cnnString = cnnStringAr(0) & sFilename
    If Dir(sFilename) = "" Then
        On Error Resume Next
        cat.Create cnnString
        If Err <> 0 Then
            CreateAccDB = False
        Else
            CreateAccDB = True
        End If
    Else
        CreateAccDB = False
    End If
    Set cat = Nothing
End Function

'输出到单元格区域
Public Function ResToSheet(ByRef rs As Variant, Optional Rng As Variant = "")
    arr = RsToArr(rs, All)
    If Not IsArray(arr) Then Exit Function
    Dim rg As Range
    If Rng = "" Then
        ActiveSheet.Range("a1").CurrentRegion.Clear
        ActiveSheet.Range("a1").Resize(UBound(arr), UBound(arr, 2)) = arr
    
    ElseIf TypeName(Rng) = "Range" Then
        Rng(1).Range("a1").CurrentRegion.Clear
        Rng(1).Range("a1").Resize(UBound(arr) + 1, UBound(arr, 2)) = arr
    Else
        Debug.Print "rng参数只能是单元格区域,或者不传参"
    End If
End Function

Private Function selProvider(dbName As String)
    If InStr(UCase(dbName), "ACCDB") > 0 Then
        sProvider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        sExtension = ""
    ElseIf InStr(UCase(dbName), "MDB") > 0 Then
        sProvider = "Provider=Microsoft.JET.OLEDB.4.0;Data Source="
        sExtension = ""
    Else
        sProvider = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="
        sExtension = ".accdb"
    End If
    selProvider = Array(sProvider, sExtension)
End Function

Private Function getProviderByExtension(fileName As String)
    Dim sExtension As String, StrRev As String, myType As cnType
    If Len(fileName) = 0 Then
        fileName = ThisWorkbook.FullName
    End If
    StrRev = StrReverse(fileName)
    sExtension = UCase(StrReverse(Split(StrRev, ".")(0)))
    Select Case sExtension
        Case "XLSX"
            myType = cnType.xlsx: Debug.Print "07+xlsx"
        Case "XLSM"
            myType = cnType.xlsm: Debug.Print "07+xlsm"
        Case "XLS"
            myType = cnType.xls: Debug.Print "03+xls"
        Case "CSV"
            myType = cnType.Csv: Debug.Print "csv"
        Case "TXT"
            myType = cnType.txt: Debug.Print "txt"
        Case "MDB"
            myType = cnType.mdb: Debug.Print "03Access"
        Case "ACCDB"
            myType = cnType.accdb: Debug.Print "07Access"
        Case Else
            Debug.Print "类型不匹配,请检查后罪名"
    End Select
    getProviderByExtension = myType
End Function


