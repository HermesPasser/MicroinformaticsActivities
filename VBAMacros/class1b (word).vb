' Activity (a)
Sub MarginNSize()
    With Application.ActiveDocument.PageSetup
		.LeftMargin = CentimetersToPoints(1.5)
		.RightMargin = CentimetersToPoints(1.5)
		.TopMargin = CentimetersToPoints(1)
		.BottomMargin = CentimetersToPoints(1)
		.PageWidth = CentimetersToPoints(21)
		.PageHeight = CentimetersToPoints(29.7)
		.Orientation = wdOrientLandscape
    End With
End Sub

' Activity (b)
Sub PageBorder()
    With Selection.Sections(1)
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdBlue 'wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdBlue
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdBlue
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleSingle
            .LineWidth = wdLineWidth050pt
            .Color = wdBlue
        End With
        With .Borders
            .DistanceFrom = wdBorderDistanceFromPageEdge
            .AlwaysInFront = True
            .SurroundHeader = True
            .SurroundFooter = True
            .JoinBorders = False
            .DistanceFromTop = 24
            .DistanceFromLeft = 24
            .DistanceFromBottom = 24
            .DistanceFromRight = 24
            .Shadow = False
            .EnableFirstPageInSection = True
            .EnableOtherPagesInSection = True
            .ApplyPageBordersToAllSections
        End With
    End With
End Sub

' CreateTable (c)
Sub CreateTable()
    ' Create table
    ActiveDocument.Tables.Add Range:=Selection.Range, NumRows:=4, NumColumns:= _
        4, DefaultTableBehavior:=wdWord9TableBehavior, AutoFitBehavior:= _
        wdAutoFitFixed

    ' Set borders
    With Selection.Tables(1)
        With .Borders(wdBorderLeft)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderRight)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderTop)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
        With .Borders(wdBorderBottom)
            .LineStyle = wdLineStyleDouble
            .LineWidth = wdLineWidth050pt
            .Color = wdColorAutomatic
        End With
    End With
    
    ' Merge cells 3x2 - 3x4
    With Selection.Tables(1)
        .Cell(row:=3, Column:=2).Merge _
        MergeTo:=.Cell(row:=3, Column:=4)
    End With

    ' Text
    Call CreateTable_addText(1, 1, 1, "Nome:")
    Call CreateTable_addText(1, 1, 3, "Rg:")
    
    Call CreateTable_addText(1, 2, 1, "Endereço:")
    Call CreateTable_addText(1, 2, 3, "Endereço:")
    
    Call CreateTable_addText(1, 3, 1, "Curso:")
    
    Call CreateTable_addText(1, 4, 1, "Semestre:")
    Call CreateTable_addText(1, 4, 3, "Turno:")
End Sub

Sub CreateTable_addText(tableIndex As Integer, row As Integer, colmn As Integer, text As String)
    Selection.Tables(tableIndex).Cell(row, colmn).Range.Font.Name = "Times New Roman"
    Selection.Tables(tableIndex).Cell(row, colmn).Range.Font.TextColor = RGB(255, 0, 0)
    Selection.Tables(tableIndex).Cell(row, colmn).Range.Font.Size = 11
    Selection.Tables(tableIndex).Cell(row, colmn).Range.text = text
End Sub
