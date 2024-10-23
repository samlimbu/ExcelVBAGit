Sub PopulateArray()
'note: will ignore blanks
   'one dimensional array
   Dim LIST() As String
   Dim strCol As Collection
   Dim rowsCount As Long
   Dim fnCall As Variant
   Dim extractColumn As String
   Dim columnNo As Integer
   Dim generatedSheetPrefix As String
   
   'config
   extractColumn = "I"
   sheetnameSource = "source"
   generatedSheetPrefix = "ws"
   filenamePrefix = "ws -"
   'config end
   
   columnNo = Range(extractColumn & 1).Column
   Debug.Print "columnNo" & columnNo
   ThisWorkbook.Activate
    
    If isWorksheetExists(sheetnameSource) Then
        Sheets(sheetnameSource).Select
    Else
        MsgBox "Sheet named source not found Using first sheet as source"
        sheetnameSource = Sheets(1).name
        Sheets(sheetnameSource).Select
    End If
    'count the rows in the range
       rowsCount = getLastRow(sheetnameSource)
       Debug.Print "rowsCount source" & rowsCount
    'run the function to create an array of unique values
       LIST() = CreateUniqueList(2, rowsCount, extractColumn)
       'loop through array
      Debug.Print Join(LIST(), vbCrLf)
       'loop through the entire array then show the element in the debug window.
     '  Dim item As Variant
      ' For Each item In LIST
         ' Debug.Print " total"; UBound(LIST); "item "; item
      ' Next item
    Debug.Print ""
    Debug.Print "populate array"; UBound(LIST)
    Dim i As Integer
    'UBound(LIST) - 1
    'For i = LBound(LIST) To UBound(LIST) - 1
    For i = LBound(LIST) To UBound(LIST) - 1
        'Debug.Print i; LIST(i)
        countRowInSheet = generateSheet(sheetnameSource, generatedSheetPrefix & i, LIST(i), columnNo)
        Debug.Print "countInSheet" & countRowInSheet - 1
        filename = filenamePrefix & " " & LIST(i) & " (" & countRowInSheet - 1 & ") " & getDate()
        fnCall = sb_Copy_Save_Worksheet_As_Workbook(filename, generatedSheetPrefix & CStr(i))
    Next
End Sub
Sub removeGeneratedSheets()
    Dim LIST() As String
    Dim extractColumn As String
    Dim rowsCount As Long
    'config
    extractColumn = "G"
    sheetnameSource = "source"
    generatedSheetPrefix = "ws"
   'config end
   
    If isWorksheetExists(sheetnameSource) Then
        Sheets(sheetnameSource).Select
    Else
        MsgBox "Sheet named source not found Using first sheet as source"
        sheetnameSource = Sheets(1).name
        Sheets(sheetnameSource).Select
    End If
   
   Sheets(sheetnameSource).Activate
    rowsCount = getLastRow(sheetnameSource)
    LIST() = CreateUniqueList(2, rowsCount, extractColumn)
     Debug.Print Join(LIST(), vbCrLf)
    For i = LBound(LIST) To UBound(LIST) - 1
    Debug.Print generatedSheetPrefix & CStr(i)
        If isWorksheetExists(generatedSheetPrefix & CStr(i)) Then
                Application.DisplayAlerts = False
                Worksheets(generatedSheetPrefix & CStr(i)).Delete
                Application.DisplayAlerts = True
        End If
    Next
    Range("A1").Select
End Sub


Function CreateUniqueList(nStart As Long, nEnd As Long, rangeStr As String) As Variant
   Dim Col As New Collection
   Dim arrTemp() As String
   Dim valCell As String
   Dim i As Integer
'Populate Temporary Collection
   On Error Resume Next
   For i = 0 To nEnd
      valCell = Range(rangeStr & nStart).Offset(i, 0).Value
      Col.Add valCell, valCell
   Next i
   Err.Clear
   On Error GoTo 0
'Resize n
   nEnd = Col.Count
'Redeclare array
   ReDim arrTemp(1 To nEnd)
'Populate temporary array by looping through the collection
   For i = 1 To Col.Count
      arrTemp(i) = Col(i)
   Next i
'return the temporary array to the function result
   CreateUniqueList = arrTemp()
End Function

Function sb_Copy_Save_Worksheet_As_Workbook(name, sheetName As String)
    Dim wb As Workbook
    Dim filename As String
    filename = "" & name & ".xlsx"
    outputFolderName = "output"
    
    'Set wb = Workbooks.Add
    ThisWorkbook.Sheets(sheetName).Copy
    
    'set output sheetname
    ActiveWorkbook.ActiveSheet.name = CStr(getDate())
    
    'no confirmation box
    Application.DisplayAlerts = False
    
    If Len(Dir(ThisWorkbook.Path & "\" & outputFolderName, vbDirectory)) = 0 Then
       MkDir ThisWorkbook.Path & "\" & outputFolderName
    End If
    
    ActiveWorkbook.SaveAs ThisWorkbook.Path & "\" & outputFolderName & "\" & filename
    ActiveWorkbook.Close SaveChanges:=False
    Application.DisplayAlerts = True
End Function

Function generateSheet(sheetnameSource, sheetName, item As String, colNo As Integer) As Variant
    
    lastRowSource = getLastRow(sheetnameSource)
    Debug.Print "lastRowSource" & lastRowSource
    ThisWorkbook.Activate

    'Sheets.Add After:=ActiveSheet
    Sheets.Add After:=Sheets(Sheets.Count)
    
    If isWorksheetExists(sheetName) Then
        Application.DisplayAlerts = False
        Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If
    
    ActiveSheet.name = sheetName
    
 
        Sheets(sheetnameSource).Select

    
    'apply filter
    ActiveSheet.Range("$A$1:$" & getLastColumnLetter() & "$" & lastRowSource).AutoFilter field:=colNo, Criteria1:=item
   
    Range("A1").Select
    
    'copy range'
    'ActiveCell.Offset(0, 0).Range("A1:" & getLastColumnLetter() & 1).Select
    Range("$A$1:$" & getLastColumnLetter() & "$" & lastRowSource).Select

    Selection.Copy

    Sheets(sheetName).Select
    
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
        
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Selection.WrapText = False
  ActiveCell.Columns("A:" & getLastColumnLetter()).EntireColumn.EntireColumn.AutoFit
    'ActiveSheet.ShowAllData
    Range("A1").Select
    generateSheet = getLastRow(ActiveSheet.name)
End Function

Function isWorksheetExists(sheetName) As Boolean
    isWorksheetExists = Evaluate("ISREF('" & sheetName & "'!A1)")
End Function

Function getLastRow(sheetName)
    If Sheets(sheetName).FilterMode Then
       Sheets(sheetName).ShowAllData
    End If
    getLastRow = Sheets(sheetName).Cells.SpecialCells(xlCellTypeLastCell).Row
End Function

Function getLastColumn()
   getLastColumn = ActiveSheet.Cells.SpecialCells(xlCellTypeLastCell).Column
End Function

'getcolumnletter
Public Function getColumnLetter(col_num)
  getColumnLetter = Split(Cells(1, col_num).Address, "$")(1)
End Function

Public Function getLastColumnLetter()
  getLastColumnLetter = Split(Cells(1, getLastColumn()).Address, "$")(1)
End Function

Function getDate()
    Dim theDate As Date
    theDate = Date
    strDate = Format(theDate, "YYYY-MM-DD")
    getDate = strDate
End Function