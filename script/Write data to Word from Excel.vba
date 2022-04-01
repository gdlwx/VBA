Sub 分离()
    Application.ScreenUpdating = False
    
    p = ThisWorkbook.Path & "/"
    f = p & "空白模板.docx"
    
    Dim myWS As Worksheet
    Set myWS = ThisWorkbook.Sheets(1) '存有数据的表格

     
    For i = 2 To 4    '遍历数据行
        FileCopy f, p & myWS.Cells(i, 3).Text & ".docx"    '复制空模板并以某列数据为名命名新产生的文档
        Set wd = CreateObject("word.application")
        Set d = wd.documents.Open(p & myWS.Cells(i, 3).Text & ".docx") '打开新文档
        
        d.InlineShapes.AddPicture Filename:=myWS.Cells(i, 6).Hyperlinks(1).Address, LinkToFile:=False, SaveWithDocument:=True
        
        d.tables(1).Cell(2, 1) = myWS.Cells(i, 3).Text '###
        '复制表格每列内容到文档，有多少项就有多少条
        d.tables(1).Cell(2, 2) = myWS.Cells(i, 5).Text '###
        d.tables(1).Cell(3, 1) = myWS.Cells(i, 12).Text '###
        
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
    f = p & "空白模板.docx"
    
    Dim myWS As Worksheet
    Set myWS = ThisWorkbook.Sheets(1) '存有数据的表格

     
    For i = 2 To 225    '遍历数据行
        FileCopy f, p & myWS.Cells(i, 1).Text & myWS.Cells(i, 4).Text & ".docx"   '复制空模板并以某列数据为名命名新产生的文档 序号+姓名.docx
        Set wd = CreateObject("word.application")
        Set d = wd.documents.Open(p & myWS.Cells(i, 1).Text & myWS.Cells(i, 4).Text & ".docx") '打开新文档
        
        d.InlineShapes.AddPicture Filename:=myWS.Cells(i, 7).Hyperlinks(1).Address, LinkToFile:=False, SaveWithDocument:=True '图片###
        
        d.tables(1).Cell(2, 1) = myWS.Cells(i, 4).Text '姓名###
        d.tables(1).Cell(2, 2) = myWS.Cells(i, 6).Text '手机号###
        d.tables(1).Cell(3, 1) = myWS.Cells(i, 13).Text '诗歌###
        
        d.Close
        wd.Quit
        Set wd = Nothing
    Next
    
    Application.ScreenUpdating = True
End Sub


'Vesion 3
'实现图片插入指定的单元格中
'参考资料：vba macro insert pictures in a word table 
'https://answers.microsoft.com/en-us/msoffice/forum/all/vba-macro-insert-pictures-in-a-word-table/09f1b75f-2879-4b21-9a53-d6c2880b6aef
' InlineShapes.AddPicture
' https://docs.microsoft.com/en-us/dotnet/api/microsoft.office.interop.word.inlineshapes.addpicture?view=word-pia


Sub 分离()
    Application.ScreenUpdating = False
    
    p = ThisWorkbook.Path & "/"
    f = p & "空白模板.docx"
    
    Dim myWS As Worksheet
    Set myWS = ThisWorkbook.Sheets(1) '存有数据的表格

     
    For i = 2 To 3    '遍历数据行
        FileCopy f, p & myWS.Cells(i, 1).Text & myWS.Cells(i, 4).Text & ".docx"   '复制空模板并以某列数据为名命名新产生的文档 序号+姓名.docx
        Set wd = CreateObject("word.application")
        Set d = wd.documents.Open(p & myWS.Cells(i, 1).Text & myWS.Cells(i, 4).Text & ".docx") '打开新文档
        
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