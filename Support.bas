Attribute VB_Name = "Support"
Public Sub RemoveButton()


    ActiveSheet.Shapes.Range(Array("Round Same Side Corner Rectangle 1")). _
        Select
    Selection.Delete
    
    'might need to move
    Columns("A:C").Select
    Selection.Delete Shift:=xlToLeft
    Columns("A:A").Select
    Selection.Insert Shift:=xlToRight
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown

End Sub


