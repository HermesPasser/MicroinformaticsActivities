VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Planilha1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Worksheet Routines, the name must be OtakuList

Private Sub CreateListLayout()
    Range("A1:F1").BorderAround 1, ColorIndex:=0, Weight:=xlThick
    Range("A1:F1").HorizontalAlignment = xlCenter
    Range("A1:F1").Font.Bold = True
    Range("A1") = "Title"
    Range("B1") = "Volume"
    Range("C1") = "Chapter"
    Range("D1") = "Status"
    Range("E1") = "Author"
    Range("F1") = "Notes"

End Sub

Private Sub CommandButton1_Click()
    Call CreateListLayout
    FormOtaku.Show
End Sub

