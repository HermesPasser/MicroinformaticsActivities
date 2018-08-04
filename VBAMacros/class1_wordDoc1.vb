' Activity (a)
Sub BorderA()
    Dim sel As Selection
    Set sel = Application.Selection
    
    If Not sel.Type = wdSelectionNormal Then
        Exit Sub ' if selection is null then close it
    End If
    
    sel.Borders.Enable = True
    sel.Borders.Shadow = True
    sel.Borders.OutsideColor = RGB(0, 0, 255)
    sel.Borders.OutsideLineStyle = wdLineStyleDouble
    sel.Borders.InsideLineWidth = wdLineWidth050pt
End Sub

' Activity (b)
Sub BorderB()
    Dim sel As Selection
    Set sel = Application.Selection
    
    If Not sel.Type = wdSelectionNormal Then
        Exit Sub ' if selection is null then close it
    End If
    
    sel.Borders.Enable = True
    sel.Borders.Shadow = True
    sel.Borders.OutsideColor = RGB(0, 100, 255)
    sel.Borders.OutsideLineStyle = wdLineStyleDoubleWavy
    sel.Borders.InsideLineWidth = wdLineWidth075pt
End Sub