Attribute VB_Name = "Player"
Option Explicit
Private Sub AddPlayer()
Dim lastRow As Integer
Dim newPlayer As String
Dim firstI As String
Dim lnIndex As Integer
Dim lastName As String



    'find end of row
   lastRow = ActiveWorkbook.Sheets("Dash").Cells(Rows.Count, "B").End(xlUp).Row
  
    'text box for name
    newPlayer = InputBox("Please enter the player's name:", "Enter the new player name", "Enter Name Here")
    
    'txtbox to B - Cell
    ActiveWorkbook.Sheets("Dash").Cells(lastRow + 1, "B").Value = newPlayer
    
    'Get last name and first inital
    firstI = Left(newPlayer, 1)
    
   lastName = GetLastName(newPlayer)
    

    
    newPlayer = firstI & " " & lastName
    
    'create tab 'copy from temp
    Worksheets("Temp").Copy After:=Worksheets("Dash")
    
    'name tab to name
    Worksheets("Temp (2)").Name = newPlayer
   
    
End Sub
Public Function GetLastName(newPlayer) As String
Dim lnIndex As Integer
    
    lnIndex = InStr(1, newPlayer, " ")
    
    GetLastName = Right(newPlayer, Len(newPlayer) - lnIndex)
End Function

Public Sub CalcData()
Dim iRow As Integer
Dim lastRow As Integer
Dim cStat As Double
Dim Ws As Worksheet
Dim sheetName As String
Dim objTable As ListObject

    ActiveSheet.Range("AI1").Value = "Calc"
    sheetName = ActiveSheet.Name
    iRow = 2

    lastRow = ActiveSheet.Cells(Rows.Count, "D").End(xlUp).Row
    
    Do Until iRow = (lastRow + 1) ' will process last time on last row, the blank row is last row + 1
        cStat = ActiveSheet.Cells(iRow, "AG").Value + ActiveSheet.Cells(iRow, "AA").Value + _
            ActiveSheet.Cells(iRow, "AB").Value + ActiveSheet.Cells(iRow, "AC").Value + _
            ActiveSheet.Cells(iRow, "AD").Value
        ActiveSheet.Cells(iRow, "AI").Value = cStat
        iRow = iRow + 1
    Loop
    
    Worksheets(sheetName).Range("D1:AI" & lastRow).Select
    Set objTable = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
    objTable.Name = sheetName
    objTable.TableStyle = "TableStyleMedium2"
    
    RemoveButton
    SortCalc (ActiveSheet.Name)

End Sub


Sub test()
    SortCalc ("K Durant")
    'SortVersus "Jordan Vs. James", GetLastName(Worksheets("Dash").ComboBox1.Value) & "_v_" & GetLastName(Worksheets("Dash").ComboBox2.Value)
End Sub

Public Sub SortCalc(sheetName As String)
Dim loName As String

    
    Sheets(sheetName).Select
        
    loName = Left(sheetName, 1) & "_" & Right(sheetName, Len(sheetName) - 2)
    
    ActiveWorkbook.Worksheets(sheetName).ListObjects(sheetName).Sort.SortFields. _
        Clear

    ActiveWorkbook.Worksheets(sheetName).ListObjects(sheetName).Sort.SortFields. _
        Add2 Key:=Range(loName & "[[#All],[Calc]]"), SortOn:=xlSortOnValues, Order _
        :=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets(sheetName).ListObjects(sheetName).Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub

