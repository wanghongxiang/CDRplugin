Sub operAllPage()

    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrMiddleRight
    
    Dim lenth As Integer
    For Each page In Application.Documents(1).Pages
   
        Debug.Print "-----------------" + page.Name
    
         ' 在页面中遍历全部的 Shape ，并将全部的组合解除掉
        For Each Shape In page.Shapes
            If Shape.Type = cdrGroupShape Then
               Dim grp As ShapeRange
               ' Shape.UngroupAllEx
               Set grp = Shape.UngroupEx
               
               ' 遍历全部的 shape 找到名称是 H_5.门牌地名区域文字 这个是批量生成的，所以名称相同
                For Each subShape In grp.Shapes
                    If subShape.Type = 6 And InStr(subShape.Name, "门牌地名区域文字") <> 0 Then
                        lenth = subShape.Text.Story.Length
                        subShape.SizeWidth = lenth * 4.2445
                        Debug.Print "=================" + CStr(subShape.Text.Story.Text)
                    End If
                Next
                
                
                ' 遍历 所有的 shape 找到名称是 H_6.编号区域文字  这个是批量生成的，所以名称相同
                ' 按照要求3位的数字，区域长度76ms  4位 97ms  6位 108ms
                ' 3位 数字和二维码之间 10ms 4位于二维码之间 7ms 6位 与二维码之间 3ms
            
                For Each subShape In grp.Shapes
                    If subShape.Type = 6 And InStr(subShape.Name, "编号区域文字") <> 0 Then
                        lenth = subShape.Text.Story.Length
                        If lenth = 3 Then
                            subShape.SizeWidth = 76
                        End If
                        
                        If lenth = 4 Then
                            subShape.SizeWidth = 97
                            subShape.Move 3, -0#
                        End If
                        
                        If lenth = 6 Then
                            subShape.SizeWidth = 108
                            subShape.Move 7, -0#
                        End If
                        
                    End If
                Next
               
               grp.Group
            End If
        Next
   Next
End Sub


Sub activePage()

    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrMiddleRight
    
    Dim lenth As Integer
    ' 在页面中遍历全部的 Shape ，并将全部的组合解除掉
    For Each Shape In ActiveDocument.activePage.Shapes
        If Shape.Type = cdrGroupShape Then
           Dim grp As ShapeRange
           ' Shape.UngroupAllEx
           Set grp = Shape.UngroupEx
           
           ' 遍历全部的 shape 找到名称是 H_5.门牌地名区域文字 这个是批量生成的，所以名称相同
            For Each subShape In grp.Shapes
                If subShape.Type = 6 And InStr(subShape.Name, "门牌地名区域文字") <> 0 Then
                    lenth = subShape.Text.Story.Length
                    subShape.SizeWidth = lenth * 4.2445
                    Debug.Print "=================" + CStr(subShape.Text.Story.Text)
                End If
            Next
            
            
            ' 遍历 所有的 shape 找到名称是 H_6.编号区域文字  这个是批量生成的，所以名称相同
            ' 按照要求3位的数字，区域长度76ms  4位 97ms  6位 108ms
            ' 3位 数字和二维码之间 10ms 4位于二维码之间 7ms 6位 与二维码之间 3ms
        
            For Each subShape In grp.Shapes
                If subShape.Type = 6 And InStr(subShape.Name, "编号区域文字") <> 0 Then
                    lenth = subShape.Text.Story.Length
                    If lenth = 3 Then
                        subShape.SizeWidth = 76
                    End If
                    
                    If lenth = 4 Then
                        subShape.SizeWidth = 97
                        subShape.Move 3, -0#
                    End If
                    
                    If lenth = 6 Then
                        subShape.SizeWidth = 108
                        subShape.Move 7, -0#
                    End If
                    
                End If
            Next
           
           grp.Group
        End If
    Next
End Sub

