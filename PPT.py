Sub RemoveEmptyTextBoxes()
    Dim SlideObj As Slide
    Dim ShapeObj As Shape
    Dim ShapeIndex As Integer
    For Each SlideObj In ActivePresentation.Slides '遍历当前ppt的每个slide
        For ShapeIndex = SlideObj.Shapes.Count To 1 Step -1 '每个slide的形状总数从n到1的逆序
            Set ShapeObj = SlideObj.Shapes(ShapeIndex) '遍历每个形状
            If ShapeObj.Type = msoTextBox Then '如果形状是文本框
                If Trim(ShapeObj.TextFrame.TextRange.Text) = "" Then '除去前后空格后，如果文本框的文字为空
                    ShapeObj.Delete '删除文本框
                End If
            End If
        Next ShapeIndex
    Next SlideObj
End Sub