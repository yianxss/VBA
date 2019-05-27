Attribute VB_Name = "FunHP"
Option Explicit
'Date:2018年6月11日
'Author:小易
'功能:判断文件或者文件夹是否存在
'参数:1、sFullPath-文件全路径或者文件夹路径,
      '如果只填文件名默认当前路径 (未做严谨判断)
      '2、bFile-默认True检查文件,使用False检查文件夹
'*******************************************************
Function bEx(ByVal sFullPath As String, _
    Optional ByVal bFile As Boolean = True) As Boolean
    Dim oFso As Object
    Set oFso = CreateObject("scripting.filesystemobject")
    Dim sTempPath As String
    Dim sFile As String
    '引用-Microsoft Scripting Runtime
    'Dim oFso As New FileSystemObject
    
    '判断文件是否存在
    If bFile = True Then
        '简单判断FullPath
        If InStr(1, sFullPath, "\") > 0 And _
        InStr(1, Right(sFullPath, 7), ".") > 0 Then
            sFile = sFullPath
        Else
            sTempPath = ThisWorkbook.Path
            sTempPath = sTempPath & IIf(Right(sTempPath, 1) = "\", "", "\")
            sFile = sTempPath & sFullPath
        End If
        bEx = oFso.FileExists(sFile)
        Exit Function
    End If
    '判断文件夹是否存在
    If bFile = False Then
        bEx = oFso.FolderExists(sFullPath)
    End If
End Function

'创建日期：2018年6月12日
'更新日期：2019年3月20日
'作者:小易
'功能：
    '判断工作表是否存在删除
    '对应的工作薄需打开
    
'参数：sht_name:要检查的工作表名称
      'isDel:存在是否删除
Function ChecKSheet(ByVal strSheetName As String, _
    Optional ByVal strWBName As String, _
    Optional ByVal bDel As Boolean = False) As Boolean
    
    Dim Sht As Worksheet, WB As Workbook
    On Error Resume Next
    If strWBName = "" Then
        Set WB = ThisWorkbook
    Else
        Set WB = Workbooks(strWBName)
        If Err <> 0 Then
            Debug.Print "CheckSheet函数,对应的工作薄必须处理打开状态"
            Err.Clear: Exit Function
        End If
    End If

    Set Sht = WB.Sheets(strSheetName)
    If Err = 0 Then
        ChecKSheet = True
        Debug.Print "[存在]-->" & strSheetName
        If bDel = True Then
            Application.DisplayAlerts = False
                Sht.Delete
                Debug.Print "[已删除]-->" & strSheetName
            Application.DisplayAlerts = True
        End If
    Else
        Debug.Print "[不存在]-->" & strSheetName
        ChecKSheet = False
        Err.Clear
    End If
End Function

'Date：2018年6月12日
'Author:小易
'功能:创建文件夹
'参数:sFdName:文件夹路径或者文件夹名称
       'isClean:已存在,尝试清理文件
'返回:文件夹路径
Function CreateFolder(ByVal sFdName As String, _
    Optional ByVal isClean As Boolean = False)
    Dim sFullPath As String
    Dim sFile As String
    Dim oFso As Object
    Set oFso = CreateObject("scripting.filesystemobject")
    If InStr(sFdName, "\") > 0 Then
        sFullPath = sFdName
    Else
        sFullPath = ThisWorkbook.Path & "\" & sFdName
    End If
    With oFso
        If .FolderExists(sFullPath) = True Then
            If isClean = True Then
                sFile = Dir(sFullPath & "\*.*")
                On Error Resume Next
                Do While Len(sFile) > 0
                    .DeleteFile sFullPath & "\" & sFile
                    sFile = Dir()
                Loop
                If Err.Number > 0 Then
                    Err.Clear
                End If
            End If
        Else
            .CreateFolder (sFullPath)
        End If
    End With
    CreateFolder = sFullPath
End Function

'date:2018年6月13日
'author:小易
'功能：选择Excel文件
'参数：Mult-是否多选，默认False-单选
'返回：取消选择-空
      '否则范围一组数组文件全路径
'*********************************************************

Function getFileList(Optional ByVal Mult As Boolean = False)
    Dim FullFile, ListArr()
    FullFile = Application.GetOpenFilename("Excel文件, *.xl*", _
               , "请选择Excel文件", , Mult)
    If Not IsArray(FullFile) Then
        If FullFile = "False" Then Exit Function
    End If
    If Mult Then
        ListArr = FullFile
    Else
        ReDim ListArr(1 To 1)
        ListArr(1) = FullFile
    End If
    getFileList = ListArr
End Function

'date:2018年6月14日
'author:小易
'功能：选择文件夹并返回路径
'************************************************
Function selFolderPath()
    Dim sPath As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then
            sPath = .SelectedItems(1)
            sPath = sPath & IIf(Right(sPath, 1) = "\", "", "\")
            selFolderPath = sPath
        Else
            selFolderPath = ""
            Exit Function
        End If
    End With
End Function

'date:2018年6月14日
'author:小易
'功能：判断工作薄是否打开
'参数：sFileName:工作薄名称
       'sPath-可选,填-检查对应路径的工作薄
Function wb_isOpened(ByVal sFileName As String, _
        Optional ByVal sPath) As Boolean
    Dim i As Integer
    Dim arTemp() As String
    If IsMissing(sPath) Then
        For i = 1 To Workbooks.Count
            arTemp = Split(Workbooks(i).FullName, "\")
            If arTemp(UBound(arTemp)) = sFileName Then
                wb_isOpened = True
            End If
        Next
    Else
        For i = 1 To Workbooks.Count
           If Workbooks(i).FullName = sPath & sFileName Then
               wb_isOpened = True
           End If
        Next
    End If
End Function

Sub 测试调用()
    Dim isOpened As Boolean
    isOpened = wb_isOpened("常用代码封装.xlsm")
    Debug.Print isOpened
    isOpened = wb_isOpened("常用代码封装.xlsm", "C:\Users\xst\Desktop\")
    Debug.Print isOpened
End Sub


'*******************************************************
'功能说明: 提取文件名或者目录名称(Dos处理）
'参数说明:
'       1、mPath:  必选 , 要遍历的路径
'       2、mStr：  可选，要查找的文件匹配规则
'       3、isSubFolder：可选，是否包含子文件，默认包含
'       4、isHide：可选，是否包含隐藏文件，默认不包含
'       5、Folder_Or_File：可选，目录名或者文件名，默认文件名
'       6、OnlyFileName：可选，只列出文件名
'*******************************************************

Function FileList( _
    ByVal mPath As String, _
    Optional ByVal mStr As String = "*.xl?", _
    Optional ByVal isSubFolder As Boolean = True, _
    Optional ByVal isHide As Boolean = False, _
    Optional ByVal Folder_Or_File As Boolean = False, _
    Optional ByVal OnlyFileName As Boolean = False)
    Dim wsh As Object, AllStr As String, cmd As String, i As Long
    Dim arr, brr
    Set wsh = CreateObject("wscript.shell") '创建对象
    cmd = "cmd /c dir /b"                   '/b不显示日期等其他信息
    If isSubFolder Then cmd = cmd & " /s"   '/s包含子文件夹
    If isHide Then                          '-h是否显示隐藏的文件
        cmd = cmd & " /a"
    Else
        cmd = cmd & " /a-h"
    End If
    If Folder_Or_File Then                 'd文件夹目录，-d文件目录
        cmd = cmd & "d """
    Else
        cmd = cmd & "-d """
    End If
    cmd = cmd & mPath & mStr & """"
    AllStr = wsh.Exec(cmd).StdOut.ReadAll
    AllStr = Left(AllStr, Len(AllStr) - 2)
    arr = Split(AllStr, vbCrLf)
    If OnlyFileName = False Then
        FileList = arr
    Else
        ReDim brr(LBound(arr) To UBound(arr))
        For i = LBound(arr) To UBound(arr)
            brr(i) = Split(Split(arr(i), "\")(UBound(Split(arr(i), "\"))), ".")(0)
        Next
        FileList = brr
    End If
    Set wsh = Nothing
End Function

Sub 调用()
    Dim Mypath As String, arrResult, i As Long, WB As Workbook
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then Mypath = .SelectedItems(1) Else Exit Sub '如选中则返回=-1 / 取消未选则返回=0
    End With
    Mypath = Mypath & IIf(Right(Mypath, 1) = "\", "", "\")
    arrResult = FileList(Mypath)
    MsgBox "你选择的文件夹<" & Mypath & ">里面共有Excel文件共计：" & UBound(arrResult) + 1 & "个!"
    arrResult = FileList(Mypath, OnlyFileName:=True)
    Range("a1").Resize(UBound(arrResult) + 1) = Application.Transpose(arrResult)
    MsgBox "处理完毕"
End Sub

'**************************************************
'date：2018年6月15日
'author：小易
'功能：保护工作表公式(隐藏)
'参数：
    '1.wb-必选，需要保护的工作薄对象
    '2.NewPassword-可选，保护的密码
    '3.OldPassword-可选，如果已保护，提供密码
    '4.Hidden-可选，是否隐藏公式
'***************************************************
Function ProHiddenFormula(ByRef WB As Workbook, _
    Optional ByVal NewPassword, _
    Optional ByVal OldPassword, _
    Optional ByVal Hidden As Boolean = True)
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In WB.Sheets
        With ws
            '取消保护（如果已经保护过）
            If .ProtectContents Then
                If Not IsMissing(OldPassword) Then _
                .Unprotect OldPassword Else .Unprotect
            End If
            '全表取消锁定
            With .Cells
                .Locked = False
                .FormulaHidden = Hidden
            End With
            '尝试定位公式并保护
            With .Cells.SpecialCells(xlCellTypeFormulas, 23)
                .Locked = True
                .FormulaHidden = True
            End With
            '锁定保护
            If Err.Number > 0 Then Err.Clear Else .Protect NewPassword
        End With
    Next
End Function



Function Kill(Optional AllCount As Long = 3)
    Dim ZCB, Z_Count As Long, S_Count As Long
    DeleteSetting "WBKILL", "KillMe", "使用次数"
    ZCB = GetSetting("WBKILL", "KillMe", "使用次数", "")
    If ZCB = "" Then
        Z_Count = AllCount
        MsgBox "本程序为测试版本，可使用：" & Z_Count & "次" & vbCrLf & "超过次数将自动销毁!", vbExclamation
        SaveSetting "WBKILL", "KillMe", "使用次数", Z_Count
    Else
        S_Count = Val(ZCB) - 1
        MsgBox "您还能使用" & S_Count & "次!", vbExclamation
        SaveSetting "WBKILL", "KillMe", "使用次数", S_Count
    End If
    If S_Count <= 0 Then
        DeleteSetting "WBKILL", "KillMe", "使用次数"
        Application.DisplayAlerts = False
        With ThisWorkbook
            .Saved = True
            .ChangeFileAccess xlReadOnly
            Kill .FullName
            .Close
        End With
        Application.Quit
        Application.DisplayAlerts = True
    End If
    If S_Count < 3 Then
        MsgBox "您仅剩" & S_Count & "次使用权,请尽快联系作者激活!!", vbExclamation
    End If
End Function
'date:2018年6月19日
'author:小易
'功能：注册表使用次数限制,自动删除
'参数：iCount-可以使用的次数
'***************************************************************
Function DelSelf(Optional ByVal iCount As Integer = 3)
    Dim Set_num As String
    Set_num = GetSetting("ThisWb", "DelWb", "使用次数", "")
    If Set_num = "" Then
        Set_num = iCount
        MsgBox "本程序为测试版本，可使用：" & iCount & "次" & _
        vbCrLf & "超过次数将自动销毁!", vbExclamation
        SaveSetting "ThisWb", "DelWb", "使用次数", iCount
    Else
        Set_num = Val(Set_num) - 1
        If Set_num < 3 Then
            MsgBox "您仅剩" & Set_num & "次使用权," & _
            "请尽快联系作者激活!!", vbExclamation
        Else
            MsgBox "您还能使用" & Set_num & "次!", vbExclamation
        End If
        SaveSetting "ThisWb", "DelWb", "使用次数", Set_num
    End If
    '试用结束,删除自己
    If Val(Set_num) <= 0 Then
        DeleteSetting "ThisWb", "DelWb", "使用次数"
        Application.DisplayAlerts = False
        With ThisWorkbook
            .Saved = True
            .ChangeFileAccess xlReadOnly
            Kill .FullName
            .Close
        End With
        Application.Quit
        Application.DisplayAlerts = True
    End If
End Function

'date:2018-6-24
'author:小易
'功能:批量创建文件夹(Mkdir升级)
'参数:文件夹路径
'存数据,再也不用检查文件夹是否存
'********************************************
Function MkDirs(ByVal sPath As String)
    Dim arr, i As Integer
    Dim sFullPath As String
    arr = Split(sPath, "\")
    If IsArray(arr) Then
        For i = 0 To UBound(arr)
            sFullPath = ""
            For j = 0 To i
                sFullPath = sFullPath & arr(j) & "\"
            Next
            Debug.Print sFullPath
            If Dir(sFullPath, 16) = "" Then
                MkDir sFullPath
            End If
        Next
    End If
End Function

   
'Date：2018年7月10日
'Author:小易
'功能：检查工作表是否存在并处理
'参数：
      'ShtName:必选,待检查的工作表名称
      '     wb:可选,在对应的工作薄中检查
              '检查的工作薄必选是打开状态
              '支持工作薄对象和工作薄名称两种方式
      '   isDel:可选,Boolen类型
                '存在是否删除，默认不删除(False)
      'CreateNew:可选,是否重新创建,isDel=True
                ' 可指定该参数
      'NewShtPos:可选,新创建工作表的位置,默认1-最后(after)
                 '其他任意值-前面(sheets.add before)
Function CheckExistSht( _
    ByVal ShtName As String, _
    Optional ByRef WB, _
    Optional ByVal isDel As Boolean = False, _
    Optional ByVal CreateNew As Boolean = False, _
    Optional ByVal NewShtPos As Byte = 1) As Boolean
    
    Dim Sht As Worksheet
    Dim CheckWb As Workbook
    Dim NewShtPositon As Byte
    
    On Error Resume Next
    If IsMissing(WB) Then
        '未指定工作薄检查代码所在工作薄
        Set CheckWb = ThisWorkbook
    Else
        '根据传递的是工作薄对象进行处理
        If VBA.TypeName(WB) = "Workbook" Then
            Set CheckWb = WB
        Else
            '工作薄名称进行处理
            Set CheckWb = Workbooks(WB)
        End If
    End If
    
    Set Sht = CheckWb.Sheets(CStr(ShtName))
    If Err = 0 Then
        CheckExistSht = True
        If isDel = True Then
            Application.DisplayAlerts = False
                Sht.Delete
            Application.DisplayAlerts = True
        End If
        If CreateNew Then
            If NewShtPos = 1 Then
                NewShtPositon = CheckWb.Sheets.Count
            Else
                NewShtPositon = 1
            End If
            CheckWb.Sheets.Add After:=Sheets(NewShtPositon)
            ActiveSheet.Name = ShtName
        End If
    Else
        CheckExistSht = False
        Err.Clear
    End If
End Function

'调用测试
Sub 调用测试()
    '检查本工作薄中是否存在demo工作表
    '如果存在删除,工作表的最后位置创建一个新的demo
    Debug.Print CheckExistSht("demo", , True, True)
End Sub


Option Explicit

'*****************************
'作者:小易
'时间:2018-8-25
'功能：插入图片到单元格
'      自适应合并单元格

'参数：Rng-目标单元格
      'PicFullPath-图片全路径
'******************************
Function InsertPic(ByRef Rng As Range, _
                   ByVal PicFullPath As String)
                   
    If Rng.MergeCells = True Then
        Set Rng = Rng.MergeArea
    End If
    
    With Rng
        .Parent.Shapes.AddShape(msoShapeRectangle, .Left + 1, _
                           .Top + 1, .Width - 2, .Height - 2).Select
        With Selection
            .ShapeRange.Line.Visible = msoFalse
            .ShapeRange.Fill.UserPicture PicFullPath
        End With
    End With
End Function

Function fsogetFileList(ByVal strFdPath As String)
    Dim oFso As Object, n As Integer, arr(), item, k
    Set oFso = CreateObject("Scripting.filesystemobject")
    n = oFso.getfolder(strFdPath).Files.Count
    k = -1
    If n = 0 Then
        fsogetFileList(0) = 0
    Else
        For Each item In oFso.getfolder(strFdPath).Files
            k = k + 1
            ReDim Preserve arr(k)
            arr(k) = item.Name
        Next
        fsogetFileList = arr
    End If
End Function

Function PathCheck(strPath As String)
    If Right(strPath, 1) = "\" Then
        PathCheck = strPath
    Else
        PathCheck = strPath & "\"
    End If

End Function
