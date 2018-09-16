VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormOtaku 
   Caption         =   "OtakuList"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5115
   OleObjectBlob   =   "FormOtaku.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FormOtaku"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form Otaku

Private Sub UpdateRowSource()
    ComboBox1.RowSource = "a2:d51"
End Sub

Private Sub UserForm_Initialize()
   Call UpdateRowSource
   Application.EnableCancelKey = xlDisabled
End Sub

Private Sub MultiPage1_Change()
    Call UpdateRowSource
End Sub

Private Sub btnDel_Click()
    If Me.ComboBox1.ListIndex = -1 Then
        MsgBox "Select a title.", , "Aviso"
        Exit Sub
    End If
    
    If MsgBox("You are sure you want remove?", vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If
    
    line = ComboBox1.ListIndex
    Range("A2").Offset(line, 0).Select
    ActiveCell.EntireRow.Delete
    Call UpdateRowSource
End Sub

Private Sub btnEdit_Click()
    If Me.ComboBox1.ListIndex = -1 Then
        MsgBox "Selecione um nome.", , "Aviso"
        Exit Sub
    End If
    
    index = ComboBox1.ListIndex

    ' For some reason i cannot pass directly  to the cell
    Vol = Me.txtEditVol.Text
    Chap = Me.txtEditChapter.Text
    Status = Me.txtEditAuthor.Text
    Author = Me.txtEditStatus.Text
    Notes = Me.txtEditNotes.Text
    
    Range("A2").Select
    ActiveCell.Offset(index, 0) = Me.txtEditTitle.Text
    ActiveCell.Offset(index, 1) = Vol
    ActiveCell.Offset(index, 2) = Chap
    ActiveCell.Offset(index, 3) = Status
    ActiveCell.Offset(index, 4) = Author
    ActiveCell.Offset(index, 5) = Notes
    
    ComboBox1.ListIndex = -1
    Me.txtEditTitle.Text = ""
    Me.txtEditVol.Text = ""
    Me.txtEditChapter.Text = ""
    Me.txtEditStatus.Text = ""
    Me.txtEditAuthor.Text = ""
    Me.txtEditNotes.Text = ""
    Call UpdateRowSource
End Sub


Private Sub ComboBox1_Change()
   ' Because the combo gets from the 2rd (A2) and the index
   ' starts with 0 (cause is an array)
    Dim index As Integer
    index = ComboBox1.ListIndex + 2
    
    Me.txtEditTitle.Text = Range("A" & index)
    Me.txtEditVol.Text = Range("B" & index)
    Me.txtEditChapter.Text = Range("C" & index)
    Me.txtEditStatus.Text = Range("D" & index)
    Me.txtEditAuthor.Text = Range("E" & index)
    Me.txtEditNotes.Text = Range("F" & index)
End Sub

Private Sub btnAdd_Click()
    If Me.txtAddTitle = "" Then
        MsgBox "A title was not inserted.", , "Error"
        Me.txtAddTitle.SetFocus
        Exit Sub
    End If
    
    Dim rng As Range, cell As Range
    Set rng = Range("A2:A51")
    For Each cell In rng
        If cell = Me.txtAddTitle Then
            MsgBox "Title already exists"
            Exit Sub
        End If
    Next cell

    Application.EnableCancelKey = xlDisabled

    ' Add in the A2
    Range("A2").Select
    ActiveCell.EntireRow.Insert
    Range("A2").HorizontalAlignment = xlLeft
    Range("A2").Font.Bold = False

	' missing something to remove the borders 
	
    ActiveCell = Me.txtAddTitle
    ActiveCell.Offset(0, 1) = Me.txtAddVol
    ActiveCell.Offset(0, 2) = Me.txtAddChapter
    ActiveCell.Offset(0, 3) = Me.txtAddStatus
    ActiveCell.Offset(0, 4) = Me.txtAddAuthor
    ActiveCell.Offset(0, 5) = Me.txtAddNotes
    
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    
    Range("A1").Select
End Sub
