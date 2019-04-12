Sub 批量插入图片批注()

' 【使用步骤】

' 1. 在电子表格所在的文件夹中建立一个名为“图片目录”（注①）的文件夹（注②）

' 2. 将待插入批注的图片（任意文件格式）放入“图片目录”

' 3. 将各张图片的文件名（不包括扩展名）重命名为与对应的目标单元格的内容 相同 的字符串（注③）

' 4. 选定全部的目标单元格

' 5. 执行宏。（完成！）


' 【注】

' ① 也可自定义 imgDir 的值为"图片目录"之外的其他名称：
              imgDir = "图片目录"

' ② 由于图片是内嵌地插入的，所以在执行宏之后可删除整个“图片目录”

' ③ 若目标单元格包含不允许出现在文件名中的特殊字符怎么办？
'    提示两种办法：
'      1. 在执行宏之后更改单元格的内容，批注不受影响
'      2. 批注属于单元格，所以能一起移动


' [VBA get image size]( https://social.msdn.microsoft.com/Forums/office/en-US/5f375529-a002-4312-a54b-b70d6d3eb6ae )
Dim objShell As Object
Dim objDir As Object
Dim objFile As Object
Dim objFileName, objFileMainName As String

fileDir = ThisWorkbook.Path & "\" & imgDir & "\"
Set objShell = CreateObject("Shell.Application")
Set objDir = objShell.Namespace(fileDir)

' [(Rough prototype)]( www.wordlm.com/Excel/jqdq/6627.html )
Dim MR As Range
For Each MR In Selection
  If Not IsEmpty(MR) Then
    MR.Select
    MR.ClearComments
    MR.AddComment
    MR.Comment.Visible = False
    MR.Comment.Text Text:=""

    ' -------- 获取图片文件 --------
    objFileMainName = fileDir & MR.Value

    ' [VBA open a file if only know part of the file name without extension name]( https://stackoverflow.com/a/2861006 )
    objFileName = Dir(objFileMainName & ".*")

    ' [VBA check if file exists]( https://stackoverflow.com/a/33771924 )
    If Dir(objFileName, vbDirectory) = "." Then
      MsgBox "未找到指定文件。请修改图片的文件名或单元格的内容，使二者相同"
      MR.ClearComments
      Exit Sub
    End If

    MR.Comment.Shape.Fill.UserPicture PictureFile:=fileDir & objFileName

    ' -------- 调整图片尺寸 --------
    Set objFile = objDir.ParseName(objFileName)

    ' [VBA extract substrings in image attributes]( https://stackoverflow.com/a/46504821 )
    size_ = objFile.ExtendedProperty("Dimensions")
    size_delimiter = InStr(size_, "x")
    width_ = Val(Mid(size_, 2, size_delimiter - 2))
    height_ = Val(Mid(size_, size_delimiter + 2, Len(size_)))

    ' [VBA get screen resolution]( https://stackoverflow.com/a/41940087 )
    'MsgBox width_ & " x " & height_ & vbCrLf & Application.UsableWidth & " x " & Application.UsableHeight

    Select Case True
      Case width_ > Application.UsableWidth
        height_ = height_ / width_ * Application.UsableWidth * 0.75
        width_ = Application.UsableWidth * 0.75
      Case height_ > Application.UsableHeight
        width_ = width_ / height_ * Application.UsableHeight * 1.15
        height_ = Application.UsableHeight * 1.15
    End Select

    MR.Comment.Shape.Width = width_
    MR.Comment.Shape.Height = height_

  End If
Next
End Sub
