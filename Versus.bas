Attribute VB_Name = "Versus"
Option Explicit
Public Sub getVersus()
'Dim playerOne As String  ComboBox1.value
'Dim playerTwo As String  ComboBox2.value
Dim sheetName As String
Dim newVersus As String
Dim objTable As ListObject
   
    
    'get veruses name
    newVersus = GetLastName(Worksheets("Dash").ComboBox1.Value) & " Vs. " & GetLastName(Worksheets("Dash").ComboBox2.Value)
    
    'create new sheet
    Worksheets("Temp").Copy After:=Worksheets("Dash")
    
    'name tab to name
    Worksheets("Temp (2)").Name = newVersus

    CopyPlayerData Worksheets("Dash").ComboBox1.Value, newVersus, 2
    CopyPlayerData Worksheets("Dash").ComboBox2.Value, newVersus, 8
    
    Worksheets(newVersus).Range("AI1").Value = "Calc"
    Worksheets(newVersus).Range("AJ1").Value = "Initals"
    
    Worksheets(newVersus).Select
    'remove cols
    Worksheets(newVersus).Range("AH:AH").Select
    Selection.Delete Shift:=xlToLeft
    
    Worksheets(newVersus).Range("AE:AF").Select
    Selection.Delete Shift:=xlToLeft
    
    Worksheets(newVersus).Range("E:Z").Select
    Selection.Delete Shift:=xlToLeft
    
    'create table
    Worksheets(newVersus).Range("D1:K13").Select
    
    Set objTable = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    objTable.Name = GetLastName(Worksheets("Dash").ComboBox1.Value) & "_v_" & GetLastName(Worksheets("Dash").ComboBox2.Value)
    objTable.TableStyle = "TableStyleMedium2"
       
    'sort rows
    SortVersus newVersus, GetLastName(Worksheets("Dash").ComboBox1.Value) & "_v_" & GetLastName(Worksheets("Dash").ComboBox2.Value)
    
    RemoveButton
    
    Worksheets(newVersus).Range("B2:I2").Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Cells.Select
    Cells.EntireColumn.AutoFit

End Sub

Private Sub SortVersus(sheetName As String, tableName As String)
        ActiveWorkbook.Worksheets(sheetName).ListObjects(tableName).Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets(sheetName).ListObjects(tableName).Sort. _
        SortFields.Add2 Key:=Columns("J:J"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(sheetName).ListObjects(tableName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Sub CopyPlayerData(playerName As String, sheetName As String, startRow As Integer)
Dim playerSheet As String
    
    playerSheet = Left(playerName, 1) & " " & GetLastName(playerName)
    'select sheet
    Worksheets(playerSheet).Select
    'sort sheet
    SortCalc (playerSheet)
    'select 6 rows
    Worksheets(playerSheet).Range("B3:AG8").Select
    'copy 6 rows
    Worksheets(playerSheet).Range("B3:AG8").Copy
    'paste 6 rows 2:7, then 8:13
    Worksheets(sheetName).Range("D" & startRow & ":AI" & startRow + 5).PasteSpecial Paste:=xlPasteValues
    'add player initals
    Worksheets(sheetName).Range("AJ" & startRow & ":AJ" & startRow + 5).Value = Left(playerName, 1) & Left(GetLastName(playerName), 1)
  
End Sub


