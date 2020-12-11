Attribute VB_Name = "Module1"
Public Function getInputDir() As String

    '処理対象月をYYYYMMの形式で指定してください。空白ならば処理日の前月になります。
    Const yyyymm As String = ""
    Dim parentPath As String
    Dim logPath As String
    Dim temp_ymd As Date
    Dim yyyy As String
    Dim mm As String
    Dim targetFolderName As String
    
    If yyyymm = "" Then
        temp_ymd = DateAdd("m", -1, Date)
        yyyy = Year(temp_ymd)
        mm = Month(temp_ymd)
        If Len(mm) = 1 Then
            mm = "0" & mm
        End If
        targetFolderName = yyyy & mm
    Else
        targetFolderName = yyyymm
    End If
    
    parentPath = Left(ActiveWorkbook.Path, InStrRev(ActiveWorkbook.Path, "¥") - 1)
    logPath = parentPath & "¥ログデータ¥" & targetFolderName
    getInputDir = logPath

End Function

Sub run()


    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Cells.Select
    Selection.ClearContents
    
    Dim LstRow  As Long
    Dim LstRow1 As Long
    Dim LstRow2 As Long
    
    Dim dstSheet As Worksheet
    Set dstSheet = ThisWorkbook.Worksheets(2)

    Dim srcBook As Workbook
    Dim srcSheet As Worksheet

    Dim buf As String
    Dim inputPath As String
    inputPath = getInputDir()
    'buf = Dir(ActiveWorkbook.Path & "¥*.csv")
    buf = Dir(inputPath & "¥*.csv")

    Dim i As Long
    i = 0
    j = 0
    Do While buf <> ""
        i = i + 1
'        Set srcBook = Workbooks.Open(ActiveWorkbook.Path + "¥" + buf)
        Set srcBook = Workbooks.Open(inputPath + "¥" + buf)
        Set srcSheet = srcBook.Worksheets(1)
        srcSheet.Select
        
        LstRow1 = srcSheet.Cells(Rows.Count, 1).End(xlUp).Row
        If i = 1 Then
            srcSheet.Range("A4:H" & LstRow1).Copy
        Else
            srcSheet.Range("A5:H" & LstRow1).Copy
        End If
        j = j + LstRow1 - 4
        
        LstRow2 = dstSheet.Cells(Rows.Count, 1).End(xlUp).Row
        dstSheet.Range("A" & LstRow2).Offset(1, 0).PasteSpecial xlPasteAll

        srcBook.Close False

        buf = Dir()
    Loop
    
    ActiveWorkbook.Worksheets("import").Sort.SortFields.Clear
    LstRow2 = dstSheet.Cells(Rows.Count, 1).End(xlUp).Row
    ActiveWorkbook.Worksheets("import").Sort.SortFields.Add Key:=Range("A3:A" & LstRow2) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("import").Sort
        .SetRange Range("A2:H" & LstRow2)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

    dstSheet.Select
    Cells.EntireColumn.AutoFit
    Columns("A:A").ColumnWidth = 15
    Range("A1").Select
    
    k = 0
    k = LstRow2 - 2

    If j = k Then
        Compare = "OK"
    Else
        Compare = "NG"
    End If
          

'
    Sheets("import2").Select
    Cells.Select
    Selection.ClearContents
    
    Sheets("import").Select
    Range("A2").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    Range("A1").Select
    
    Sheets("import2").Select
    Cells.Select
    ActiveSheet.Paste
    Range("A1").Select

    Columns("B:B").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlFixedWidth, _
        FieldInfo:=Array(Array(0, 1), Array(10, 1)), TrailingMinusNumbers:=True
    Selection.NumberFormatLocal = "yyyy/m/d"
    Columns("B:B").Select
    Selection.NumberFormatLocal = "h:mm:ss;@"
    Cells.EntireColumn.AutoFit
    Range("A1").Select

    Columns("B:C").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Selection.NumberFormat = "General"

    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=TEXT(RC[-1],""ddd"")"
    Range("C2").Select
    ActiveCell.FormulaR1C1 = "=IF(COUNTIF(holiday!C[-2],'import2'!RC[-2])<>0,""Hol"","""")"
    Range("B2:C2").Select
    Selection.Copy
    LstRow = Sheets("import2").Cells(Rows.Count, 1).End(xlUp).Row
    Range("B2:C" & LstRow).Select
    ActiveSheet.Paste
    
    Columns("B:C").Select
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("B:C").ColumnWidth = 3
    Columns("D:D").ColumnWidth = 8
    
    Range("A1").Select
    
   
    Sheets("list").Select
    Cells.Select
    Selection.ClearContents

    Sheets("import2").Select
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.AutoFilter Field:=10, Criteria1:="<>"
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.Copy
    
    Sheets("list").Select
    Cells.Select
    ActiveSheet.Paste
    
    Sheets("import2").Select
    Selection.AutoFilter
    Range("A1").Select

    Sheets("list").Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    
    Range("L2").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(OR(RC[-10]=""Sat"",RC[-10]=""Sun"",RC[-9]=""Hol"",HOUR(RC[-8])<5,HOUR(RC[-8])>=22),"""",1)"
    Range("L2").Select
    Selection.Copy
    LstRow = Sheets("list").Cells(Rows.Count, 1).End(xlUp).Row
    Range("L2:L" & LstRow).Select
    ActiveSheet.Paste

    Columns("L:L").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.AutoFilter Field:=12, Criteria1:="="
    NGLog = WorksheetFunction.Subtotal(3, Columns(1)) - 1
    Range("A1").Select


    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
        
    ActiveWorkbook.Save

    MsgBox "CSV file count          " & i & vbCrLf & _
           "CSV import record count " & j & vbCrLf & _
           "paste record count      " & k & vbCrLf & _
           "OK or NG                " & Compare & vbCrLf & _
           "list record             " & NGLog

End Sub




