'Method 1: Merge Selected Documents into One


Sub MergeMultiDocsIntoOne()
  Dim dlgFile As FileDialog
  Dim nTotalFiles As Integer
  Dim nEachSelectedFile As Integer

  Set dlgFile = Application.FileDialog(msoFileDialogFilePicker)
 
  With dlgFile
    .AllowMultiSelect = True
    If .Show <> -1 Then
      Exit Sub
    Else
      nTotalFiles = .SelectedItems.Count
    End If
  End With
 
  For nEachSelectedFile = 1 To nTotalFiles
    Selection.InsertFile dlgFile.SelectedItems.Item(nEachSelectedFile)
    If nEachSelectedFile < nTotalFiles Then
      Selection.InsertBreak Type:=wdPageBreak
    Else
      If nEachSelectedFile = nTotalFiles Then
        Exit Sub
      End If
    End If
  Next nEachSelectedFile
End Sub



'Method 2: Merge All Documents in a Folder into One
Sub MergeFilesInAFolderIntoOneDoc()
  Dim dlgFile As FileDialog
  Dim objDoc As Document, objNewDoc As Document
  Dim StrFolder As String, strFile As String
 
  Set dlgFile = Application.FileDialog(msoFileDialogFolderPicker)
 
  With dlgFile
    If .Show = -1 Then
      StrFolder = .SelectedItems(1) & "\"
    Else
      MsgBox ("No folder is selected!")
      Exit Sub
    End If
  End With
 
  strFile = Dir(StrFolder & "*.docx", vbNormal)
  Set objNewDoc = Documents.Add
 
  While strFile <> ""
    Set objDoc = Documents.Open(FileName:=StrFolder & strFile)
    objDoc.Range.Copy
    objNewDoc.Activate
    With Selection
      .Paste
      .InsertBreak Type:=wdPageBreak
      .Collapse wdCollapseEnd
    End With
 
    objDoc.Close SaveChanges:=wdDoNotSaveChanges
 
    strFile = Dir()
  Wend
 
  objNewDoc.Activate
  Selection.EndKey Unit:=wdStory
  Selection.Delete
End Sub

'参考资料：
'https://www.datanumen.com/blogs/2-ways-quickly-merge-multiple-word-documents-one-via-vba/