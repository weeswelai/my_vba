Sub select_download_file()
    'ctrl + shift + q
    '计算文件
    Dim fso, file_name


    Set fso = CreateObject("scripting.filesystemobject") '创建并引用文件系统对象

    file_name = Application.GetOpenFilename("ALL Files(*.*),*.*")  '调用文件窗口

    If file_name <> "False" Then
        Set file_ob = fso.GetFile(file_name)
        file_date_creatd = Left(file_ob.DateCreated, InStr(1, file_ob.DateCreated, " ") - 1)  '创建日期
        file_size = file_ob.Size                '文件大小(字节)
        '字节转Kb Mb Gb
        If  (file_size == 0) Then
            file_size = "0"
        ElseIf (file_size > 0 And file_size < 1024) Then  
            file_size = "Kb"
        ElseIf (file_size >= 1024 And file_size < 1048576) Then     'Kb
            file_size = (file_size / 1024) + 0.1
            file_size2 = "Kb"
        ElseIf (file_size >= 1048576 And file_size <= 1073741824) Then 'Mb
            file_size = file_size / 1048576 + 0.1
            file_size2 = "Mb"
        ElseIf (file_size >= 1073741824) Then  
            file_size = file_size / 1073741824 + 0.1
            file_size2 = "Gb"
        End If
        file_size = Str(file_size)
        file_size = Replace(Left(file_size, InStr(1, file_size, ".") + 1), " ", "") & file_size2    '去掉小数点之后的数字，加上单位
        file_path = file_ob.ParentFolder.Path  '父文件夹
        file_type = UCase(Right(file_name, Len(file_name) - InStrRev(file_name, ".")))  '获取文件后缀
        If (file_type = "7Z" Or file_type = "RAR" Or file_type = "ZIP" Or file_type = "TAR" Or file_type = "TZ" Or file_type = "GZ") Then
            file_type = "压缩包" & file_type
        End If
        row_num = Selection.Row         '判断当前excel选中的行数 返回数字
        sheet_name = ActiveSheet.Name   '获取当前sheet页的名称
        Sheets(sheet_name).Cells(row_num, 1) = file_date_creatd  '日期
        Sheets(sheet_name).Cells(row_num, 2) = file_ob.Name  '文件名
        Sheets(sheet_name).Cells(row_num, 4) = file_type  '文件类型
        Sheets(sheet_name).Cells(row_num, 5) = file_size  '文件大小
        Sheets(sheet_name).Cells(row_num, 6) = file_path   '目录
    Else
        End
    End If

End Sub