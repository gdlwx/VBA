Option Explicit
 
Function GetImageSize(ImagePath As String) As Variant
 
    '--------------------------------------------------------------------------------------
    'Returns an array of integers that hold the image width and height in pixels.
    'The first element of the array corresponds to the width and the second to the height.
 
    'The function uses the Microsoft Windows Image Acquisition Library v2.0, which can be
    'found in the path: C:\Windows\System32\wiaaut.dll
    'However, the code is written in late binding, so no reference is required.
 
    'Written By:    Christos Samaras
    'Date:          18/02/2018
    'E-mail:        xristos.samaras@gmail.com
    'Site:          http://www.myengineeringworld.net
    '--------------------------------------------------------------------------------------
 
    'Declaring the necessary variables.
    Dim imgSize(1)  As Integer
    Dim wia         As Object
 
    'Check that the image file exists.
    If FileExists(ImagePath) = False Then Exit Function
 
    'Check that the image file corresponds to an image format.
    If IsValidImageFormat(ImagePath) = False Then Exit Function
 
    'Create the ImageFile object and check if it exists.
    On Error Resume Next
    Set wia = CreateObject("WIA.ImageFile")
    If wia Is Nothing Then Exit Function
    On Error GoTo 0
 
    'Load the ImageFile object with the specified File.
    wia.LoadFile ImagePath
 
    'Get the necessary properties.
    imgSize(0) = wia.Width
    imgSize(1) = wia.Height
 
    'Release the ImageFile object.
    Set wia = Nothing
 
    'Return the array.
    GetImageSize = imgSize
 
End Function
 
Function FileExists(FilePath As String) As Boolean
 
    '--------------------------------------------------
    'Checks if a file exists (using the Dir function).
    '--------------------------------------------------
 
    On Error Resume Next
    If Len(FilePath) > 0 Then
        If Not Dir(FilePath, vbDirectory) = vbNullString Then FileExists = True
    End If
    On Error GoTo 0
 
End Function
 
Function IsValidImageFormat(FilePath As String) As Boolean
 
    '----------------------------------------------
    'Checks if a given path is a valid image file.
    '----------------------------------------------
 
    'Declaring the necessary variables.
    Dim imageFormats    As Variant
    Dim i               As Integer
 
    'Some common image extentions.
    imageFormats = Array(".bmp", ".jpg", ".gif", ".tif", ".png")
 
    'Loop through all the extentions and check if the path contains one of them.
    For i = LBound(imageFormats) To UBound(imageFormats)
        'If the file path contains the extension return true.
        If InStr(1, UCase(FilePath), UCase(imageFormats(i)), vbTextCompare) > 0 Then
            IsValidImageFormat = True
            Exit Function
        End If
    Next i
 
End Function