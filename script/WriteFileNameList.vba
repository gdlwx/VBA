Sub writeFileNameList()
    
    Dim f As String
    MyPath = ActiveDocument.Path
    f = MyPath & "\filelist.txt"
    MyName = Dir(MyPath & "\" & "*.docx")
    
    Open f For Output As #1
     I = 0
    
    Do While MyName <> ""
        If MyName <> ActiveDocument.Name And MyName <> "all.docx" Then
          I = I + 1
          Print #1, MyName
       End If
    MyName = Dir
    Loop
   
   Close #1

End Sub
