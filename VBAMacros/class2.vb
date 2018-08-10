' Ex 1
Sub MessageFromInput()
    Dim str As String
    str = InputBox("Type something")
    Dim msg As String
    msg = "Your message is " + Chr$(13) + "'" + str + "'?"
    
    If MsgBox(msg, vbYesNo + vbExclamation, "Confirmation") = vbYes Then
        Selection.Font.Name = "Times New Roman"
        Selection.Font.Size = 16
        Selection.Font.Bold = True
        Selection.Font.Italic = True
        Selection.Font.TextColor = RGB(0, 255, 0)
        Selection.TypeText text:=str
        
        Selection.Font.Size = 12
        Selection.Font.TextColor = RGB(0, 0, 0)
        Selection.Font.Bold = False
        Selection.Font.Italic = False
    End If
End Sub

' Ex 2
Sub TypeHeader()
    Dim str As String
    str = InputBox("Type the header text:", "Header Text")
    
    ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageHeader
    
    If Not str = "" Then
        Selection.TypeText text:=str
    Else
        Selection.TypeText text:="Cabeçalho de Documento - Para Aula Prg Micro"
    End If
    
    ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
End Sub

' Ex 3
Sub ShowCalc()
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    
    Shell ("C:/windows/system32/calc.exe")
    
    MsgBox ("Press OK after calculating. Use 'ctrl + c' to copy")
    clipboard.GetFromClipboard
    Selection.TypeText text:="Calculated value: " + clipboard.GetText(1)
End Sub
