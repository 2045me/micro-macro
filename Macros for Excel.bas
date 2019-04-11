Sub 批量插入图片批注()

' 【使用步骤】
' 1. 在电子表格所在的文件夹中建立一个名为“图片目录”的文件夹
' 2. 将待插入批注的图片放入“图片目录”
' 3. 将各张图片的文件名重命名为与对应的目标单元格的内容*相同*的字符串
' 4. 选定全部的目标单元格
' 5. 执行宏

' P.S. 由于图片是内嵌地插入的，所以在执行宏之后可删除整个“图片目录”


'On Error Resume Next
Dim MR As Range

' [VBA get image size]( https://social.msdn.microsoft.com/Forums/office/en-US/5f375529-a002-4312-a54b-b70d6d3eb6ae )
Dim objShell As Object
Dim objFolder As Object
Dim objFile As Object
Dim objFileName, objFileMainName As String

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(ActiveWorkbook.Path & "\图片目录")


' [(Rough prototype)]( www.wordlm.com/Excel/jqdq/6627.html )
For Each MR In Selection
  If Not IsEmpty(MR) Then
    MR.Select
    MR.AddComment
    MR.Comment.Visible = False
    MR.Comment.Text Text:=""
    
    
    objFileMainName = ActiveWorkbook.Path & "\图片目录\" & MR.Value
    nameExtension = ".png"
    objFileName = objFileMainName & nameExtension
    
    ' [VBA check if file exists]( https://stackoverflow.com/a/33771924 )
    If Dir(objFileName, vbDirectory) = vbNullString Then
        nameExtension = ".jpg"
        objFileName = objFileMainName & nameExtension
    End If
    
    MR.Comment.Shape.Fill.UserPicture PictureFile:=objFileName
    
    
    Set objFile = objFolder.ParseName(MR.Value & nameExtension)
    
    ' [VBA extract substrings in image attributes]( https://stackoverflow.com/a/46504821 )
    size_ = objFile.ExtendedProperty("Dimensions")
    size_delimiter = InStr(size_, "x")
    width_ = Mid(size_, 2, size_delimiter - 2)
    height_ = Mid(size_, size_delimiter + 2, Len(size_) - 8)
    
    If (height_ > 500) Then
        height_ = height_ / 2
        width_ = width_ / 2
    End If
    
    MR.Comment.Shape.width = width_
    MR.Comment.Shape.height = height_
    
  End If
Next
End Sub
