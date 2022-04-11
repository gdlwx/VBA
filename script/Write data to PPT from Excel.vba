'Version 1



Sub 分离()
    Application.ScreenUpdating = False
    
    p = ThisWorkbook.Path & "/"
    f = p & "空白模板.pptx"
    
    Dim myWS As Worksheet
    Set myWS = ThisWorkbook.Sheets(1) '存有数据的表格

     
    For i = 2 To 3    '遍历数据行
        FileCopy f, p & myWS.Cells(i, 1).Text & myWS.Cells(i, 4).Text & ".pptx"   '复制空模板并以某列数据为名命名新产生的文档 序号+姓名.docx
        Set pt = CreateObject("powerpoint.application")
        Set d = pt.documents.Open(p & myWS.Cells(i, 1).Text & myWS.Cells(i, 4).Text & ".pptx") '打开新文档
        
        d.InlineShapes.AddPicture Filename:=myWS.Cells(i, 7).Hyperlinks(1).Address, LinkToFile:=False, SaveWithDocument:=True, Range:=d.tables(1).Cell(3, 1).Range '图片， 可以指定表格某单元格写入###
        
        '姓名### d.tables(1).Cell(2, 1) = myWS.Cells(i, 4).Text
        '手机号### d.tables(1).Cell(2, 2) = myWS.Cells(i, 6).Text
        '诗歌### d.tables(1).Cell(3, 1) = myWS.Cells(i, 13).Text
        
        d.Close
        wd.Quit
        Set wd = Nothing
    Next
    
    Application.ScreenUpdating = True
End Sub

'Version 2 
Sub 分离()
    Application.ScreenUpdating = False
    
    p = ThisWorkbook.Path & "/"
    f = p & "空白模板.pptx"
    
    Dim myWS As Worksheet
    Set myWS = ThisWorkbook.Sheets(1) '存有数据的表格

     
    For i = 2 To 2    '遍历数据行
        FileCopy f, p & myWS.Cells(i, 1).Text & myWS.Cells(i, 4).Text & ".pptx"   '复制空模板并以某列数据为名命名新产生的文档 序号+姓名.docx
        Set pt = CreateObject("powerpoint.application")
        Set d = pt.Presentations.Open(p & myWS.Cells(i, 1).Text & myWS.Cells(i, 4).Text & ".pptx") '打开新文档
        
        d.Slides(1).Shapes.AddPicture Filename:="C:\Users\Paul\Desktop\temp\Test\test.jpeg", LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=100, Top:=100   '图片###
        
        'd.Slides(1).Shapes.Table.Cell(1, 1).Shape.AddPicture  Filename:=myWS.Cells(i, 7).Hyperlinks(1).Address, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue '
        '姓名### d.tables(1).Cell(2, 1) = myWS.Cells(i, 4).Text
        '手机号### d.tables(1).Cell(2, 2) = myWS.Cells(i, 6).Text
        '诗歌### d.tables(1).Cell(3, 1) = myWS.Cells(i, 13).Text
        
        d.Close
        pt.Quit
        Set pt = Nothing
    Next
    
    Application.ScreenUpdating = True
End Sub


'Version 3

Sub ReadDataFromExcel()

    'Application.ScreenUpdating = False
    
    Dim myWS As Object
    Set myWS = GetObject(, "Excel.Application") '打开存有数据的表格
    
    Dim myPPT As Presentation
    Set myPPT = ActivePresentation
    
    Dim pptSlide As Slide
    
    Dim pptLayout As CustomLayout
    Set pptLayout = myPPT.Slides(1).CustomLayout

     
    For i = 2 To 4    '遍历数据行
        
        Set pptSlide = myPPT.Slides.AddSlide(i, pptLayout)
        
        Set d = myPPT.Slides(i - 1)
        
        d.Shapes.AddPicture FileName:=myWS.Sheets(1).Cells(i, 7).Hyperlinks(1).Address, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=100, Top:=100  '图片OK###
        
        d.Shapes(1).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 4).Text
        d.Shapes(1).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 6).Text
        d.Shapes(1).Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 13).Text
        

        '姓名### d.tables(1).Cell(2, 1) = myWS.Cells(i, 4).Text
        '手机号### d.tables(1).Cell(2, 2) = myWS.Cells(i, 6).Text
        '诗歌### d.tables(1).Cell(3, 1) = myWS.Cells(i, 13).Text
        
    Next
    
    'Application.ScreenUpdating = True
End Sub


'Version 4

Sub ReadDataFromExcel()

    
    Dim myWS As Object
    Set myWS = GetObject(, "Excel.Application") '打开存有数据的表格
    
    Dim myPPT As Presentation
    Set myPPT = ActivePresentation
    
    Dim pptSlide As Slide
 
    
    Dim pptLayout As CustomLayout
    Set pptLayout = myPPT.Slides(1).CustomLayout

     
    For i = 2 To 4    '遍历数据行
        
        Set pptSlide = myPPT.Slides.AddSlide(i, pptLayout)
        
        Set d = myPPT.Slides(i - 1)
        
              
        d.Shapes.AddPicture FileName:=myWS.Sheets(1).Cells(i, 7).Hyperlinks(1).Address, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=100, Top:=100  '图片OK###
        
        d.Select
        d.Shapes.AddTable(2, 3).Select
        d.Shapes(2).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(2).Table.Cell(2, 2)
        d.Shapes(2).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(2).Table.Cell(2, 3)
        
        d.Shapes(2).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 1).Text '序号
        d.Shapes(2).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 4).Text '姓名
        d.Shapes(2).Table.Cell(1, 3).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 6).Text '手机号
        d.Shapes(2).Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 13).Text '诗歌
        
    Next
    

End Sub


'Table.Cell method (PowerPoint) https://docs.microsoft.com/en-us/office/vba/api/powerpoint.table.cell
'Invalid request To select a shape, its view must be active  https://answers.microsoft.com/en-us/msoffice/forum/all/active-window-in-powerpoint/6ddd07f6-c8d3-4bde-823d-c89cd2c9a106


'Version 5

Sub ReadDataFromExcel()

    
    Dim myWS As Object
    Set myWS = GetObject(, "Excel.Application") '打开存有数据的表格
    
    Dim myPPT As Presentation
    Set myPPT = ActivePresentation
    
    Dim pptSlide As Slide
    Dim myPic As Shape
    
    Dim pptLayout As CustomLayout
    Set pptLayout = myPPT.Slides(1).CustomLayout

     
    For i = 2 To 4    '遍历数据行
        
        Set pptSlide = myPPT.Slides.AddSlide(i, pptLayout)
        
        Set d = myPPT.Slides(i - 1)
        
              
        Set myPic = d.Shapes.AddPicture(FileName:=myWS.Sheets(1).Cells(i, 7).Hyperlinks(1).Address, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=100, Top:=100)  '图片OK###
        
        With myPic
            .Scaleheight 1, msoTrue
            .Scalewidth 1, msoTrue
        End With

        d.Select
        d.Shapes.AddTable(2, 3).Select
        d.Shapes(2).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(2).Table.Cell(2, 2)
        d.Shapes(2).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(2).Table.Cell(2, 3)
        
        d.Shapes(2).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 1).Text '序号
        d.Shapes(2).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 4).Text '姓名
        d.Shapes(2).Table.Cell(1, 3).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 6).Text '手机号
        d.Shapes(2).Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 13).Text '诗歌
        
    Next
    

End Sub


'version 6 
'use objWIA to get image size of local picture.

Sub ReadDataFromExcel()

    Dim objWIA As Object

    
    Dim myWS As Object
    Set myWS = GetObject(, "Excel.Application") '打开存有数据的表格
    
    Dim myPPT As Presentation
    Set myPPT = ActivePresentation
    
    Dim pptSlide As Slide
    Dim myPic As Shape
    
    Dim pptLayout As CustomLayout
    Set pptLayout = myPPT.Slides(1).CustomLayout

     
    For i = 2 To 2    '遍历数据行
        
        Set pptSlide = myPPT.Slides.AddSlide(i, pptLayout)
        
        Set d = myPPT.Slides(i - 1)
        
        
        Set objWIA = CreateObject("WIA.ImageFile")
        objWIA.LoadFile "D:\Dev\lab\ppt\test.jpeg"
        
        Debug.Print "宽度：" & objWIA.Width & "像素"
        Debug.Print "高度：" & objWIA.Height & "像素"
        
        
      
        
              
        Set myPic = d.Shapes.AddPicture(FileName:=myWS.Sheets(1).Cells(i, 7).Hyperlinks(1).Address, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=100, Top:=100)  '图片OK###
        

        If objWIA.Width >= objWIA.Height Then
        
            With myPic
                .LockAspectRatio = msoFalse
                .Width = 800
                .Height = objWIA.Height * 800 / objWIA.Width
                .Top = 100
            End With
            
        Else
            With myPic
                .LockAspectRatio = msoFalse
                .Width = objWIA.Width * 800 / objWIA.Height
                .Height = 800
                .Top = 5
            End With
        
        End If
        
        d.Select
        d.Shapes.AddTable(2, 3).Select
        d.Shapes(2).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(2).Table.Cell(2, 2)
        d.Shapes(2).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(2).Table.Cell(2, 3)
        
        d.Shapes(2).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 1).Text '序号
        d.Shapes(2).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 4).Text '姓名
        d.Shapes(2).Table.Cell(1, 3).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 6).Text '手机号
        d.Shapes(2).Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 13).Text '诗歌
        
    Next
    

End Sub


'Version 7

Sub ReadDataFromExcel()
    
    Dim myWS As Object
    Set myWS = GetObject(, "Excel.Application") '打开存有数据的表格
    
    Dim myPPT As Presentation
    Set myPPT = ActivePresentation
    
    Dim pptSlide As Slide
    Dim myPic As Shape
    
    Dim pptLayout As CustomLayout
    Set pptLayout = myPPT.Slides(1).CustomLayout

     
    For i = 2 To 4    '遍历数据行
        
        Set pptSlide = myPPT.Slides.AddSlide(i, pptLayout)
        
        Set d = myPPT.Slides(i - 1)
        
             
        Set myPic = d.Shapes.AddPicture(FileName:=myWS.Sheets(1).Cells(i, 7).Hyperlinks(1).Address, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=0, Top:=0)  '图片OK###
        
        
        
        'myPic.LockAspectRatio = msoFalse
        'myPic.ScaleWidth 1, msoTrue
        'myPic.ScaleHeight 1, msoTrue
        If myPic.Width > 900 Then
        
            With myPic
                .LockAspectRatio = msoFalse
                .ScaleHeight 900 / myPic.Width, msoTrue
                .ScaleWidth 900 / myPic.Width, msoTrue
                .Top = 50
                .Left = 5
            End With
            
        End If

        
        d.Select
        d.Shapes.AddTable(2, 3).Select
        d.Shapes(1).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(1).Table.Cell(2, 2)
        d.Shapes(1).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(1).Table.Cell(2, 3)
        
        d.Shapes(1).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 1).Text '序号
        d.Shapes(1).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 4).Text '姓名
        d.Shapes(1).Table.Cell(1, 3).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 6).Text '手机号
        d.Shapes(1).Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 13).Text '诗歌
        
    Next
    

End Sub


'version 8


Sub ReadDataFromExcel()
    
    Dim myWS As Object
    Set myWS = GetObject(, "Excel.Application") '打开存有数据的表格
    
    Dim myPPT As Presentation
    Set myPPT = ActivePresentation
    
    Dim pptSlide As Slide
    Dim myPic As Shape
    
    Dim pptLayout As CustomLayout
    Set pptLayout = myPPT.Slides(1).CustomLayout

     
    For i = 2 To 4    '遍历数据行
        
        Set pptSlide = myPPT.Slides.AddSlide(i, pptLayout)
        
        Set d = myPPT.Slides(i - 1)
        
             
        Set myPic = d.Shapes.AddPicture(FileName:=myWS.Sheets(1).Cells(i, 7).Hyperlinks(1).Address, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=0, Top:=0)  '图片OK###
        
        
        
        'myPic.LockAspectRatio = msoFalse
        'myPic.ScaleWidth 1, msoTrue
        'myPic.ScaleHeight 1, msoTrue
        If myPic.Width > 900 Then
        
            With myPic
                .LockAspectRatio = msoFalse
                .ScaleHeight 900 / myPic.Width, msoTrue
                .ScaleWidth 900 / myPic.Width, msoTrue
                .Top = 50
                .Left = 5
            End With
            
        End If

        
        d.Select
        d.Shapes.AddTable(2, 3).Select
        d.Shapes(1).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(1).Table.Cell(2, 2)
        d.Shapes(1).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(1).Table.Cell(2, 3)
        
        d.Shapes(1).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 1).Text '序号
        d.Shapes(1).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 4).Text '姓名
        d.Shapes(1).Table.Cell(1, 3).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 6).Text '手机号
        d.Shapes(1).Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 13).Text '诗歌
        
    Next
    

End Sub


'Version 9

Sub ReadDataFromExcel()
    
    Dim myWS As Object
    Set myWS = GetObject(, "Excel.Application") '打开存有数据的表格
    
    Dim myPPT As Presentation
    Set myPPT = ActivePresentation
    
    p = myPPT.Path & "\"
    
    Dim pptSlide As Slide
    Dim myPic As Shape
    
    Dim pptLayout As CustomLayout
    Set pptLayout = myPPT.Slides(1).CustomLayout
    
    Set fso = CreateObject("Scripting.FileSystemObject")

     
    For i = 2 To 101    '遍历数据行
        
        Set pptSlide = myPPT.Slides.AddSlide(i, pptLayout)
        
        Set d = myPPT.Slides(i - 1)
        
        picName = myWS.Cells(i, 6)
        ext = fso.GetExtensionName(picName)
        picPath = p & myWS.Cells(i, 1) & myWS.Cells(i, 4) & "." & ext
        
        If InStr(1, "jpg jpeg png gif", ext) > 0 Then
             
            Set myPic = d.Shapes.AddPicture(FileName:=picPath, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, Left:=0, Top:=0)  '图片OK###
             
        
             If myPic.Width > 900 Then
        
                   With myPic
                        .LockAspectRatio = msoFalse
                        .ScaleHeight 900 / myPic.Width, msoTrue
                        .ScaleWidth 900 / myPic.Width, msoTrue
                        .Top = 50
                        .Left = 5
                    End With
            
            End If

        
            d.Select
            d.Shapes.AddTable(2, 3).Select
            d.Shapes(1).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(1).Table.Cell(2, 2)
            d.Shapes(1).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(1).Table.Cell(2, 3)
        
            d.Shapes(1).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 1).Text '序号
            d.Shapes(1).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 4).Text '姓名
            d.Shapes(1).Table.Cell(1, 3).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 5).Text '手机号
            d.Shapes(1).Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 7).Text '诗歌
        Else
            d.Select
            d.Shapes.AddTable(2, 4).Select
            d.Shapes(1).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(1).Table.Cell(2, 2)
            d.Shapes(1).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(1).Table.Cell(2, 3)
            d.Shapes(1).Table.Cell(2, 1).Merge MergeTo:=d.Shapes(1).Table.Cell(2, 4)
        
            d.Shapes(1).Table.Cell(1, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 1).Text '序号
            d.Shapes(1).Table.Cell(1, 2).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 4).Text '姓名
            d.Shapes(1).Table.Cell(1, 3).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 5).Text '手机号
            d.Shapes(1).Table.Cell(1, 4).Shape.TextFrame.TextRange.Text = "视频文件"
            d.Shapes(1).Table.Cell(2, 1).Shape.TextFrame.TextRange.Text = myWS.Sheets(1).Cells(i, 7).Text '诗歌
        End If
        
    Next

End Sub