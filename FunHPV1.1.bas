Attribute VB_Name = "FunHP"
Option Explicit
'Date:2018��6��11��
'Author:С��
'����:�ж��ļ������ļ����Ƿ����
'����:1��sFullPath-�ļ�ȫ·�������ļ���·��,
      '���ֻ���ļ���Ĭ�ϵ�ǰ·�� (δ���Ͻ��ж�)
      '2��bFile-Ĭ��True����ļ�,ʹ��False����ļ���
'*******************************************************
Function bEx(ByVal sFullPath As String, _
    Optional ByVal bFile As Boolean = True) As Boolean
    Dim oFso As Object
    Set oFso = CreateObject("scripting.filesystemobject")
    Dim sTempPath As String
    Dim sFile As String
    '����-Microsoft Scripting Runtime
    'Dim oFso As New FileSystemObject
    
    '�ж��ļ��Ƿ����
    If bFile = True Then
        '���ж�FullPath
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
    '�ж��ļ����Ƿ����
    If bFile = False Then
        bEx = oFso.FolderExists(sFullPath)
    End If
End Function

'�������ڣ�2018��6��12��
'�������ڣ�2019��3��20��
'����:С��
'���ܣ�
    '�жϹ������Ƿ����ɾ��
    '��Ӧ�Ĺ��������
    
'������sht_name:Ҫ���Ĺ���������
      'isDel:�����Ƿ�ɾ��
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
            Debug.Print "CheckSheet����,��Ӧ�Ĺ��������봦���״̬"
            Err.Clear: Exit Function
        End If
    End If

    Set Sht = WB.Sheets(strSheetName)
    If Err = 0 Then
        ChecKSheet = True
        Debug.Print "[����]-->" & strSheetName
        If bDel = True Then
            Application.DisplayAlerts = False
                Sht.Delete
                Debug.Print "[��ɾ��]-->" & strSheetName
            Application.DisplayAlerts = True
        End If
    Else
        Debug.Print "[������]-->" & strSheetName
        ChecKSheet = False
        Err.Clear
    End If
End Function

'Date��2018��6��12��
'Author:С��
'����:�����ļ���
'����:sFdName:�ļ���·�������ļ�������
       'isClean:�Ѵ���,���������ļ�
'����:�ļ���·��
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

'date:2018��6��13��
'author:С��
'���ܣ�ѡ��Excel�ļ�
'������Mult-�Ƿ��ѡ��Ĭ��False-��ѡ
'���أ�ȡ��ѡ��-��
      '����Χһ�������ļ�ȫ·��
'*********************************************************

Function getFileList(Optional ByVal Mult As Boolean = False)
    Dim FullFile, ListArr()
    FullFile = Application.GetOpenFilename("Excel�ļ�, *.xl*", _
               , "��ѡ��Excel�ļ�", , Mult)
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

'date:2018��6��14��
'author:С��
'���ܣ�ѡ���ļ��в�����·��
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

'date:2018��6��14��
'author:С��
'���ܣ��жϹ������Ƿ��
'������sFileName:����������
       'sPath-��ѡ,��-����Ӧ·���Ĺ�����
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

Sub ���Ե���()
    Dim isOpened As Boolean
    isOpened = wb_isOpened("���ô����װ.xlsm")
    Debug.Print isOpened
    isOpened = wb_isOpened("���ô����װ.xlsm", "C:\Users\xst\Desktop\")
    Debug.Print isOpened
End Sub


'*******************************************************
'����˵��: ��ȡ�ļ�������Ŀ¼����(Dos����
'����˵��:
'       1��mPath:  ��ѡ , Ҫ������·��
'       2��mStr��  ��ѡ��Ҫ���ҵ��ļ�ƥ�����
'       3��isSubFolder����ѡ���Ƿ�������ļ���Ĭ�ϰ���
'       4��isHide����ѡ���Ƿ���������ļ���Ĭ�ϲ�����
'       5��Folder_Or_File����ѡ��Ŀ¼�������ļ�����Ĭ���ļ���
'       6��OnlyFileName����ѡ��ֻ�г��ļ���
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
    Set wsh = CreateObject("wscript.shell") '��������
    cmd = "cmd /c dir /b"                   '/b����ʾ���ڵ�������Ϣ
    If isSubFolder Then cmd = cmd & " /s"   '/s�������ļ���
    If isHide Then                          '-h�Ƿ���ʾ���ص��ļ�
        cmd = cmd & " /a"
    Else
        cmd = cmd & " /a-h"
    End If
    If Folder_Or_File Then                 'd�ļ���Ŀ¼��-d�ļ�Ŀ¼
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

Sub ����()
    Dim Mypath As String, arrResult, i As Long, WB As Workbook
    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show Then Mypath = .SelectedItems(1) Else Exit Sub '��ѡ���򷵻�=-1 / ȡ��δѡ�򷵻�=0
    End With
    Mypath = Mypath & IIf(Right(Mypath, 1) = "\", "", "\")
    arrResult = FileList(Mypath)
    MsgBox "��ѡ����ļ���<" & Mypath & ">���湲��Excel�ļ����ƣ�" & UBound(arrResult) + 1 & "��!"
    arrResult = FileList(Mypath, OnlyFileName:=True)
    Range("a1").Resize(UBound(arrResult) + 1) = Application.Transpose(arrResult)
    MsgBox "�������"
End Sub

'**************************************************
'date��2018��6��15��
'author��С��
'���ܣ�����������ʽ(����)
'������
    '1.wb-��ѡ����Ҫ�����Ĺ���������
    '2.NewPassword-��ѡ������������
    '3.OldPassword-��ѡ������ѱ������ṩ����
    '4.Hidden-��ѡ���Ƿ����ع�ʽ
'***************************************************
Function ProHiddenFormula(ByRef WB As Workbook, _
    Optional ByVal NewPassword, _
    Optional ByVal OldPassword, _
    Optional ByVal Hidden As Boolean = True)
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In WB.Sheets
        With ws
            'ȡ������������Ѿ���������
            If .ProtectContents Then
                If Not IsMissing(OldPassword) Then _
                .Unprotect OldPassword Else .Unprotect
            End If
            'ȫ��ȡ������
            With .Cells
                .Locked = False
                .FormulaHidden = Hidden
            End With
            '���Զ�λ��ʽ������
            With .Cells.SpecialCells(xlCellTypeFormulas, 23)
                .Locked = True
                .FormulaHidden = True
            End With
            '��������
            If Err.Number > 0 Then Err.Clear Else .Protect NewPassword
        End With
    Next
End Function



Function Kill(Optional AllCount As Long = 3)
    Dim ZCB, Z_Count As Long, S_Count As Long
    DeleteSetting "WBKILL", "KillMe", "ʹ�ô���"
    ZCB = GetSetting("WBKILL", "KillMe", "ʹ�ô���", "")
    If ZCB = "" Then
        Z_Count = AllCount
        MsgBox "������Ϊ���԰汾����ʹ�ã�" & Z_Count & "��" & vbCrLf & "�����������Զ�����!", vbExclamation
        SaveSetting "WBKILL", "KillMe", "ʹ�ô���", Z_Count
    Else
        S_Count = Val(ZCB) - 1
        MsgBox "������ʹ��" & S_Count & "��!", vbExclamation
        SaveSetting "WBKILL", "KillMe", "ʹ�ô���", S_Count
    End If
    If S_Count <= 0 Then
        DeleteSetting "WBKILL", "KillMe", "ʹ�ô���"
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
        MsgBox "����ʣ" & S_Count & "��ʹ��Ȩ,�뾡����ϵ���߼���!!", vbExclamation
    End If
End Function
'date:2018��6��19��
'author:С��
'���ܣ�ע���ʹ�ô�������,�Զ�ɾ��
'������iCount-����ʹ�õĴ���
'***************************************************************
Function DelSelf(Optional ByVal iCount As Integer = 3)
    Dim Set_num As String
    Set_num = GetSetting("ThisWb", "DelWb", "ʹ�ô���", "")
    If Set_num = "" Then
        Set_num = iCount
        MsgBox "������Ϊ���԰汾����ʹ�ã�" & iCount & "��" & _
        vbCrLf & "�����������Զ�����!", vbExclamation
        SaveSetting "ThisWb", "DelWb", "ʹ�ô���", iCount
    Else
        Set_num = Val(Set_num) - 1
        If Set_num < 3 Then
            MsgBox "����ʣ" & Set_num & "��ʹ��Ȩ," & _
            "�뾡����ϵ���߼���!!", vbExclamation
        Else
            MsgBox "������ʹ��" & Set_num & "��!", vbExclamation
        End If
        SaveSetting "ThisWb", "DelWb", "ʹ�ô���", Set_num
    End If
    '���ý���,ɾ���Լ�
    If Val(Set_num) <= 0 Then
        DeleteSetting "ThisWb", "DelWb", "ʹ�ô���"
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
'author:С��
'����:���������ļ���(Mkdir����)
'����:�ļ���·��
'������,��Ҳ���ü���ļ����Ƿ��
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

   
'Date��2018��7��10��
'Author:С��
'���ܣ���鹤�����Ƿ���ڲ�����
'������
      'ShtName:��ѡ,�����Ĺ���������
      '     wb:��ѡ,�ڶ�Ӧ�Ĺ������м��
              '���Ĺ�������ѡ�Ǵ�״̬
              '֧�ֹ���������͹������������ַ�ʽ
      '   isDel:��ѡ,Boolen����
                '�����Ƿ�ɾ����Ĭ�ϲ�ɾ��(False)
      'CreateNew:��ѡ,�Ƿ����´���,isDel=True
                ' ��ָ���ò���
      'NewShtPos:��ѡ,�´����������λ��,Ĭ��1-���(after)
                 '��������ֵ-ǰ��(sheets.add before)
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
        'δָ�����������������ڹ�����
        Set CheckWb = ThisWorkbook
    Else
        '���ݴ��ݵ��ǹ�����������д���
        If VBA.TypeName(WB) = "Workbook" Then
            Set CheckWb = WB
        Else
            '���������ƽ��д���
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

'���ò���
Sub ���ò���()
    '��鱾���������Ƿ����demo������
    '�������ɾ��,����������λ�ô���һ���µ�demo
    Debug.Print CheckExistSht("demo", , True, True)
End Sub


Option Explicit

'*****************************
'����:С��
'ʱ��:2018-8-25
'���ܣ�����ͼƬ����Ԫ��
'      ����Ӧ�ϲ���Ԫ��

'������Rng-Ŀ�굥Ԫ��
      'PicFullPath-ͼƬȫ·��
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
