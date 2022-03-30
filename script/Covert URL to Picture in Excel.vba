Sub 把F列图片链接在K列生成图片()
    '以第2-4行为例，设定单位格宽度（图片宽度）
    Columns("K:K").ColumnWidth = 36.22
    Rows("2:209").RowHeight = 150
    '插入图片（假设图都是4：3的，否则还涉及要取得原始尺寸，如果图小，可以直接把200，150改为-1，-1）
    For i = 2 To 209
        Cells(i, 7) = Cells(i, 6).Hyperlinks(1).Address
        ActiveSheet.Shapes.AddPicture(Cells(i, 7).Value, msoCTrue, msoCTrue, Range("K" & i).Left, Range("K" & i).Top, 200, 150).AlternativeText = "Error"
    Next
End Sub
