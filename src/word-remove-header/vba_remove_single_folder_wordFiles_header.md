# 文件批量处理

- 使用nodejs去掉A下文件名的 ”-备战“后面的文字，存储到B
- 然后复制B到B - 副本
- 使用VBA宏批量去掉页眉页脚

  - 打开一个word
  - ALT + F11
  - 插入 -> 模块
  - 粘贴代码
  - F5 运行
  - 选择文件夹路径
```vba
Sub RemoveHeadersAndFootersFromFolder()
    Dim folderPath As String
    Dim fileName As String
    Dim doc As Document
    Dim section As section
    Dim header As headerFooter
    Dim footer As headerFooter
    ' 弹出文件夹选择框
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "请选择文件夹"
        If .Show = -1 Then ' 如果选择了文件夹
            folderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "没有选择文件夹，操作已取消。"
            Exit Sub
        End If
    End With
    ' 获取文件夹中的所有 .docx 文件
    fileName = Dir(folderPath & "*.docx")
    ' 遍历文件夹中的每个文件
    Do While fileName <> ""
        ' 打开文件
        Set doc = Documents.Open(folderPath & fileName)
        ' 如果文档受保护，解除保护
        If doc.ProtectionType <> wdNoProtection Then
            doc.Unprotect ' 如果文档受保护，解除保护
        End If
        ' 遍历文档中的所有节（section）并删除页眉和页脚
        For Each section In doc.Sections
            ' 确保断开与前一节的链接（如果存在）
            If section.Headers.Count > 0 Then
                section.Headers(1).LinkToPrevious = False
            End If
            If section.Footers.Count > 0 Then
                section.Footers(1).LinkToPrevious = False
            End If
            ' 删除页眉（所有类型）
            On Error Resume Next
            For Each header In section.Headers
                header.Range.Delete
            Next header
            On Error GoTo 0
            ' 删除页脚（所有类型）
            On Error Resume Next
            For Each footer In section.Footers
                footer.Range.Delete
            Next footer
            On Error GoTo 0
        Next section
        ' 保存并关闭文档
        doc.Save
        doc.Close
        ' 处理下一个文件
        fileName = Dir
    Loop

    MsgBox "所有文件处理完成！"
End Sub
```