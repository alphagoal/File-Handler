
Sub TestFunc()

'ImportOneCsvSheet (11)
'ImportOneExcel (8)

'Dim result As Variant
'result = FileOpenable(16)
'MsgBox result

ExportOneCsv (24)

'Dim result As Variant
'result = CheckTabExistInCsvGen(Cells(12, 3))
'MsgBox result


End Sub

Sub ScanFolder()

Dim foldpath As String
Dim filepath1 As String
Dim filename As String
Dim filetype As String
Dim i        As Integer


clear_control_input
foldpath = Range("FolderPath").Value                                'Subject to change in future if also scan other files inside the folder
If Right(foldpath, 1) <> "\" Then foldpath = foldpath & "\"
filetype = LCase(Range("FileType").Value)
If filetype <> "" Then If Left(filetype, 1) <> "." Then filetype = "." & filetype
If filetype = "" Then filepath1 = foldpath Else: filepath1 = foldpath & "*" & filetype
filename = Dir(filepath1)

If filename = "" Then
    MsgBox "Folder/ File does not exist!"
    Exit Sub
End If

i = 1
While filename <> ""
    Cells(Range("Col_FoldPath").Row + i, Range("Col_FoldPath").Column) = foldpath
    Cells(Range("Col_FileName").Row + i, Range("Col_FileName").Column) = filename
    Cells(Range("Col_FullPath").Row + i, Range("Col_FullPath").Column) = foldpath & filename
    Cells(Range("Col_FileType").Row + i, Range("Col_FileType").Column) = FindFileType(Range("Col_FileType").Row + i, Range("Col_FullPath").Column)
    i = i + 1
    filename = Dir
Wend

DynamicIndex


End Sub

Sub ImportAllTabs()

Dim start_row, end_row, i As Integer
    start_row = Range("Col_FullPath").Row + 1
    end_row = Range("Col_FullPath").End(xlDown).Row

Dim startcell_1, endcell_1, startcell_2, endcell_2, startcell_3, endcell_3 As String
    startcell_1 = Cells(start_row, Range("Col_Imp_Time").Column).Address
    endcell_1 = Cells(1048000, Range("Col_Imp_Time").Column).Address
    startcell_2 = Cells(start_row, Range("Col_Remarks").Column).Address
    endcell_2 = Cells(1048000, Range("Col_Remarks").Column).Address
    'startcell_3 = Cells(start_row, Range("Col_Exp_Time").Column).Address
    'endcell_3 = Cells(1048000, Range("Col_Exp_Time").Column).Address
    Range(startcell_1 & ":" & endcell_1).ClearContents
    Range(startcell_2 & ":" & endcell_2).ClearContents
    'Range(startcell_3 & ":" & endcell_3).ClearContents

For i = start_row To end_row
    If FileOpenable(i) = False Then
        If Cells(i, Range("Col_Imp_YN").Column) = "Y" Or Cells(i, Range("Col_Imp_YN").Column) = "y" Then
            Dim ErrMsg3 As String
            ErrMsg3 = "Import Error 03: File cannot be opened by Excel"
            MsgBox ErrMsg3
            Cells(i, Range("Col_Remarks").Column) = ErrMsg3
        End If
    Else:
        If Cells(i, Range("Col_FileType").Column) = "csv" Then
            If Cells(i, Range("Col_Imp_YN").Column) = "Y" Or Cells(i, Range("Col_Imp_YN").Column) = "y" Then ImportOneCsvSheet (i)
        Else:
            If Cells(i, Range("Col_Imp_YN").Column) = "Y" Or Cells(i, Range("Col_Imp_YN").Column) = "y" Then ImportOneExcel (i)
        End If
    End If
Next i

End Sub

Sub ExportAll()

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.AskToUpdateLinks = False

Dim start_row, end_row, i As Integer
    start_row = Range("Col_FullPath").Row + 1
    end_row = Range("Col_FullPath").End(xlDown).Row

Dim startcell_1, endcell_1, startcell_2, endcell_2, startcell_3, endcell_3 As String
    'startcell_1 = Cells(start_row, Range("Col_Imp_Time").Column).Address
    'endcell_1 = Cells(1048000, Range("Col_Imp_Time").Column).Address
    startcell_2 = Cells(start_row, Range("Col_Remarks").Column).Address
    endcell_2 = Cells(1048000, Range("Col_Remarks").Column).Address
    startcell_3 = Cells(start_row, Range("Col_Exp_Time").Column).Address
    endcell_3 = Cells(1048000, Range("Col_Exp_Time").Column).Address
    'Range(startcell_1 & ":" & endcell_1).ClearContents
    Range(startcell_2 & ":" & endcell_2).ClearContents
    Range(startcell_3 & ":" & endcell_3).ClearContents

For i = start_row To end_row
    If Cells(i, Range("Col_Exp_YN").Column) = "Y" Or Cells(i, Range("Col_Exp_YN").Column) = "y" Then ExportOneCsv (i)
Next i


End Sub

Function CheckTabExistInCsvGen(tabname As String)

Dim sht_check As Worksheet
Dim tabexist As Boolean
tabexist = False

For Each sht_check In ThisWorkbook.Worksheets
    If sht_check.Name = tabname Then
        tabexist = True
        Exit For
    End If
Next sht_check

CheckTabExistInCsvGen = tabexist

End Function

Sub ExportOneCsv(rownum As Integer)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.AskToUpdateLinks = False

Dim tabname, ExpPath As String
    ExpPath = Cells(rownum, Range("Col_ExpPath").Column)
    If ExpPath <> "" And ExpPath <> "." And ExpPath <> ".." And Len(ExpPath) >= 2 Then
        'MsgBox Right(ExpPath, 1)
        If Right(ExpPath, 1) <> "\" Then ExpPath = ExpPath & "\"
    End If
    tabname = Cells(rownum, Range("Col_TabName").Column)
Dim FolderExists As String
    FolderExists = Dir(ExpPath, vbDirectory)
Dim OrgMsg As String
    OrgMsg = Cells(rownum, Range("Col_Remarks").Column)

    If ExpPath = "" Or ExpPath = "." Or ExpPath = ".." Or Len(ExpPath) < 2 Or FolderExists = "" Then
        Dim ErrMsg4 As String
        ErrMsg4 = "Export Error 04: Export path does not exist. Please correct"
        If OrgMsg = "" Then Cells(rownum, Range("Col_Remarks").Column) = ErrMsg4 Else: Cells(rownum, Range("Col_Remarks").Column) = OrgMsg & vbNewLine & ErrMsg4
        MsgBox ErrMsg4
        Exit Sub
    End If

If CheckTabExistInCsvGen(Cells(rownum, Range("Col_TabName").Column)) = False Then
    Dim ErrMsg5 As String
    ErrMsg5 = "Export Error 05: Tab name for CSV export does not exist in this tool. Please correct"
    If OrgMsg = "" Then Cells(rownum, Range("Col_Remarks").Column) = ErrMsg5 Else: Cells(rownum, Range("Col_Remarks").Column) = OrgMsg & vbNewLine & ErrMsg5
    MsgBox ErrMsg5
    Exit Sub
End If

Dim ExpRange As String
ExpRange = FindBorderRange(tabname)
'MsgBox ExpRange
If ExpRange = "Error06" Then
    ErrMsg6 = "Export Error 06: Cannot Define the Export Range"
    Cells(rownum, Range("Col_Remarks").Column) = ErrMsg6
    Exit Sub
Else:
    Cells(rownum, Range("Col_Exp_Range").Column) = ExpRange
End If


Sheets(tabname).Range(ExpRange).Copy
Workbooks.Add
Set wb = activeworkbook
Range("A1").PasteSpecial xlPasteValues
Application.CutCopyMode = True

Application.DisplayAlerts = False
wb.SaveAs filename:=ExpPath & tabname, FileFormat:=xlCSV, CreateBackup:=False
wb.Close

Sheets("Control").Select
Cells(rownum, Range("Col_Exp_Time").Column).Value = Format(Now, "dd-mmm-yyyy h:mm:ss AM/PM") 'Column M

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.AskToUpdateLinks = True

End Sub

Sub ImportOneCsvSheet(rownum As Integer)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.AskToUpdateLinks = False

Dim SourceBook As Workbook
Dim folderpath, filename, fullpath As String
    folderpath = Cells(rownum, Range("Col_FoldPath").Column) 'Column D
    filename = Cells(rownum, Range("Col_FileName").Column) 'Column E
    fullpath = Cells(rownum, Range("Col_FullPath").Column) 'Column F
    Cells(rownum, 1).Select

Dim index As Variant
Dim Tabame, TempTabName, ImpRange As String
    index = Cells(ActiveCell.Row, Range("Col_index").Column)
    tabname = Sheets("Control").Cells(rownum, Range("Col_TabName").Column) 'Column C
    If tabname = "" Then
        tabname = "index" & "_" & index
        Sheets("Control").Cells(rownum, Range("Col_TabName").Column) = tabname
    End If
    ImpRange = Cells(rownum, Range("Col_Imp_Range").Column) 'Column K

Dim sht_check As Worksheet
For Each sht_check In ThisWorkbook.Worksheets
    If sht_check.Name = tabname Then
        sht_check.Delete
        Exit For
    End If
Next sht_check

Application.ScreenUpdating = False
Select Case ImpRange

    Case ""
        Application.ScreenUpdating = False
        Set SourceBook = Workbooks.Open(fullpath)
        SourceBook.Sheets(1).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ActiveSheet.Name = tabname
        ActiveSheet.Rows(1).Insert
        ActiveSheet.Columns(1).Insert
        Columns.EntireColumn.Hidden = False
        SourceBook.Close

    Case Else
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = tabname
                
        Dim CopyRange As Range
        Dim DestinationRange As Range
        Set SourceBook = Workbooks.Open(fullpath)
        Set CopyRange = SourceBook.Sheets(1).Range(ImpRange)
        Set DestinationRange = ThisWorkbook.Worksheets(tabname).Range("A1")
        CopyRange.Copy DestinationRange
        SourceBook.Close
        ActiveSheet.Rows(1).Insert
        ActiveSheet.Columns(1).Insert
        Columns.EntireColumn.Hidden = False
        
End Select

'clear_marker
BorderArd

Sheets("Control").Select
Cells(rownum, Range("Col_Imp_Time").Column).Value = Format(Now, "dd-mmm-yyyy h:mm:ss AM/PM") 'Column J

Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.AskToUpdateLinks = True

End Sub

Sub ImportOneExcel(rownum As Integer)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.AskToUpdateLinks = False

Sheets("Control").Select

Dim ErrorRange As Range
Set ErrorRange = ThisWorkbook.Worksheets("Control").Cells(rownum, Range("Col_Remarks").Column)
Dim thiswb As Workbook
Set thiswb = ThisWorkbook

Dim SourceBook As Workbook
Dim folderpath, filename, fullpath As String
    folderpath = Cells(rownum, Range("Col_FoldPath").Column) 'Column D
    filename = Cells(rownum, Range("Col_FileName").Column) 'Column E
    fullpath = Cells(rownum, Range("Col_FullPath").Column) 'Column F
    Cells(rownum, 1).Select

Dim index As Variant
Dim Tabame, ImpRange, ImpTab, ErrorText  As String
    index = Cells(ActiveCell.Row, Range("Col_index").Column)
    tabname = Sheets("Control").Cells(rownum, Range("Col_TabName").Column) 'Column C
    If tabname = "" Then
        tabname = "index" & "_" & index
        Sheets("Control").Cells(rownum, Range("Col_TabName").Column) = tabname
    End If
    ImpRange = Cells(rownum, Range("Col_Imp_Range").Column) 'Column K
    ImpTab = Cells(rownum, Range("Col_Imp_Tab").Column) 'Column L

Dim sht_check As Worksheet
For Each sht_check In ThisWorkbook.Worksheets
    If sht_check.Name = tabname Then
        sht_check.Delete
        Exit For
    End If
Next sht_check

 
If ImpTab = "" Then
    ErrorText = "Import Error 01: Missing Import Tab Name. Please input correct tab name."
    Cells(rownum, Range("Col_Remarks").Column) = ErrorText
    MsgBox ErrorText
    Exit Sub
Else:
    'May add error handling in future, e.g ". on erorr goto..." for broken excel files
    Application.ScreenUpdating = False
    Set SourceBook = Workbooks.Open(fullpath)
    Dim CheckTabExist As Boolean
    CheckTabExist = False
    Dim sht As Worksheet
    For Each sht In SourceBook.Worksheets
        If LCase(ImpTab) = LCase(sht.Name) Then
            'MsgBox "Worksheet '" & ImpTab & "' exists!"
            CheckTabExist = True
            Exit For
        End If
    Next sht
    
    If CheckTabExist = False Then
        SourceBook.Close
        ErrorText = "Import Error 02: Import tab name does not exist in the source workbook. Please correct the name."
        ErrorRange = ErrorText
        MsgBox ErrorText
        'Cells(rownum, Range("Col_Remarks").Column) = ErrorText
        Exit Sub
    End If
End If


Select Case ImpRange

    Case ""
        Application.ScreenUpdating = False
        'Set SourceBook = Workbooks.Open(fullpath)
        thiswb.Activate
        thiswb.Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = tabname
        SourceBook.Sheets(ImpTab).Cells.Copy
        thiswb.Worksheets(tabname).Range("A1").PasteSpecial xlPasteValues
        'thiswb.Worksheets(tabname).Range("A1").PasteSpecial xlPasteFormats
        
        SourceBook.Close
        ActiveSheet.Name = tabname
        ActiveSheet.Rows(1).Insert
        ActiveSheet.Columns(1).Insert
        Columns.EntireColumn.Hidden = False
        Rows.EntireRow.Hidden = False
        

    Case Else
        thiswb.Activate
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.Name = tabname
                
        Dim CopyRange As Range
        Dim DestinationRange As Range
        'Set SourceBook = Workbooks.Open(fullpath)
        Set CopyRange = SourceBook.Sheets(ImpTab).Range(ImpRange)
        Set DestinationRange = ThisWorkbook.Worksheets(tabname).Range("A1")
        CopyRange.Copy
        DestinationRange.PasteSpecial xlPasteValues
        'DestinationRange.PasteSpecial xlPasteFormats
        SourceBook.Close
        ActiveSheet.Rows(1).Insert
        ActiveSheet.Columns(1).Insert
        Columns.EntireRow.Hidden = False
End Select

'clear_marker
BorderArd

Sheets("Control").Select
Cells(rownum, Range("Col_Imp_Time").Column).Value = Format(Now, "dd-mmm-yyyy h:mm:ss AM/PM") 'Column J


Application.ScreenUpdating = True
Application.DisplayAlerts = True
Application.EnableEvents = True
Application.AskToUpdateLinks = True

End Sub

Sub testborderfunc()

Dim result As Variant
'Dim rng As Range
'Set rng = Range("E11")

result = testborder_color(Cells(15, 5))

MsgBox result

End Sub

Function test_leftborder_validcolor(check_cell As Range)

Dim colorindex As Variant
Dim result As Boolean
colorindex = check_cell.Borders(xlEdgeLeft).Color
test_leftborder_validcolor = colorindex = 16711680  'color index of blue

'Columns(3).Borders(xlEdgeRight).LineStyle = xlContinuous
End Function

Function test_rightborder_validcolor(check_cell As Range)

Dim colorindex As Variant
Dim result As Boolean
colorindex = check_cell.Borders(xlEdgeRight).Color
test_rightborder_validcolor = colorindex = 16711680  'color index of blue

End Function

Sub test_findborderrange()

Dim testtabname, result As String

'testtabname = Cells(29, 3).Value
'MsgBox testtabname
'testtabname = Cells(23 + Range("Col_TabName").Row, Range("Col_TabName").Column).Value

result = FindBorderRange(Cells(29, 3))

'result = FindBorderRange(testtabname)
MsgBox result

End Sub


Function FindBorderRange(tabname As Variant)

Application.DisplayAlerts = False
Application.EnableEvents = False
Application.ScreenUpdating = False
Application.AskToUpdateLinks = False

On Error GoTo Err

Dim firstcol, lastcol, firstrow, lastrow, col As Long
Dim sheetlastrow, sheetlastcol As Long

Sheets(tabname).Select
sheetlastrow = Range("A:A").Rows.Count
sheetlastcol = Range("1:1").Columns.Count
'MsgBox sheetlastrow & " x " & sheetlastcol

Dim found_firstborder_col As Boolean
found_firstborder_col = False

col = 0

Dim temprow As Long
Dim hasborder_left As Boolean
hasborder_left = False

While found_firstborder_col = False
    col = col + 1
    If WorksheetFunction.CountA(Columns(col)) > 0 Then
        temprow = 1
        While temprow < sheetlastrow And found_firstborder_col = False
            hasborder_left = test_leftborder_validcolor(Cells(temprow, col))
            If hasborder_left = True Then found_firstborder_col = True Else: temprow = Cells(temprow, col).End(xlDown).Row
        Wend
    End If
Wend

firstcol = col
'MsgBox "The first column number is: " & firstcol


Dim found_lastborder_col As Boolean
found_lastborder_col = False
Dim hasborder_right As Boolean
hasborder_right = False

col = col - 1

While found_lastborder_col = False
    col = col + 1
    If WorksheetFunction.CountA(Columns(col)) > 0 Then
        temprow = 1
        While temprow < sheetlastrow And found_lastborder_col = False
            hasborder_right = test_rightborder_validcolor(Cells(temprow, col))
            If hasborder_right = True Then found_lastborder_col = True Else: temprow = Cells(temprow, col).End(xlDown).Row
        Wend
    End If
Wend

lastcol = col
'MsgBox "The last column number is: " & lastcol


Dim startrow_address, endrow_address As String
Dim marker_startrow, marker_endrow As Range
Set marker_startrow = ActiveSheet.Cells.Find("MarkStartRow")
Set marker_endrow = ActiveSheet.Cells.Find("MarkEndRow")

firstrow = marker_startrow.Row
lastrow = marker_endrow.Row
'MsgBox "The first row is: " & firstrow & vbNewLine & "The last row is: " & lastrow

Dim firstcell, lastcell As Range
Set firstcell = Cells(firstrow, firstcol)
Set lastcell = Cells(lastrow, lastcol)

Dim ExportRange As String
ExportRange = firstcell.Address & ":" & lastcell.Address
'MsgBox "The Export Range is: " & ExportRange

Sheets("Control").Select

'Application.DisplayAlerts = True
'Application.EnableEvents = True
'Application.ScreenUpdating = True
'Application.AskToUpdateLinks = True

FindBorderRange = ExportRange

Exit Function

Err:
    Dim ErrMsg6 As String
    ErrMsg6 = "Export Error 06: Cannot Define the Export Range"
    MsgBox ErrMsg6
    Sheets("Control").Select
    
    FindBorderRange = "Error06"
    
'Application.DisplayAlerts = True
'Application.EnableEvents = True
'Application.ScreenUpdating = True
'Application.AskToUpdateLinks = True
    
    Exit Function

End Function

Sub BorderArd()

Dim nocells As Long
nocells = WorksheetFunction.CountA(Range(Cells(1, 1), Cells(Range("a1").End(xlDown).Row, Range("a1").End(xlToRight).Column)))
'MsgBox nocells

Dim col, firstcol, lastcol As Long
col = 1

Dim foundfirstcol As Boolean
foundfirstcol = False
While foundfirstcol = False
    If WorksheetFunction.CountA(Columns(col)) > 0 Then foundfirstcol = True
    col = col + 1
Wend
col = col - 1
firstcol = col


Dim countedcells, col_nocells As Long
countedcells = 0
While countedcells < nocells
    col_nocells = WorksheetFunction.CountA(Columns(col))
    countedcells = countedcells + col_nocells
    col = col + 1
Wend
col = col - 1
lastcol = col


Dim temprow1, temprow2, rowsearch, sheetlastrow, toprow, botrow As Long
Dim i As Long
sheetlastrow = Range("a1").End(xlDown).Row
toprow = sheetlastrow
botrow = 1
rowsearch = 1

For i = firstcol To lastcol
    temprow1 = Cells(1, i).End(xlDown).Row
    If temprow1 < toprow Then toprow = temprow1
    temprow2 = Cells(sheetlastrow, i).End(xlUp).Row
    If temprow2 > botrow Then botrow = temprow2
Next i

Range(Cells(toprow, firstcol), Cells(botrow, lastcol)).BorderAround _
    LineStyle:=xlContinuous, _
    Weight:=xlThick, _
    Color:=vbBlue


Cells(toprow, lastcol + 1) = "MarkStartRow"
Cells(toprow, lastcol + 1).Font.Color = vbRed
Cells(toprow, lastcol + 1).Interior.Color = vbYellow
Cells(toprow, lastcol + 1).AddComment ("Do not change this marker")
With Cells(toprow, lastcol + 1).Comment.Shape.TextFrame.Characters.Font
    .Size = 16
    .Name = "Arial Black"
End With

Cells(botrow, lastcol + 1) = "MarkEndRow"
Cells(botrow, lastcol + 1).Font.Color = vbRed
Cells(botrow, lastcol + 1).Interior.Color = vbYellow
Cells(botrow, lastcol + 1).AddComment ("Do not change this marker")
With Cells(botrow, lastcol + 1).Comment.Shape.TextFrame.Characters.Font
    .Size = 16
    .Name = "Arial Black"
End With


End Sub

Sub clear_marker()

Dim startrow_address, endrow_address As String
Dim marker_startrow, marker_endrow As Range
Set marker_startrow = ActiveSheet.Cells.Find("MarkStartRow")
Set marker_endrow = ActiveSheet.Cells.Find("MarkEndRow")


On Error GoTo Err
startrow_address = marker_startrow.Address
endrow_address = marker_endrow.Address
'MsgBox startrow_address & vbNewLine & endrow_address

Range(startrow_address).Clear
Range(endrow_address).Clear
'
Err:
    'MsgBox "marker NOT cleared"
    Exit Sub
    Resume Next

End Sub


Function FindFileType(i As Integer, j As Integer)

Dim temparry() As String
Dim last_arryno As Integer
temparry = Split(Cells(i, j).Value, ".")
last_arryno = UBound(temparry)
FindFileType = temparry(last_arryno)

End Function

Sub DynamicIndex()

Dim i, lastindex As Integer
lastindex = Range("Col_FullPath").End(xlDown).Row - Range("Col_FullPath").Row

For i = 1 To lastindex
    Range("Col_index").Offset(i, 0).Value = i
Next i

End Sub

Function FileOpenable(rownum As Integer)

Dim filetype() As Variant
'must dim as variant

filetype() = Array("csv", "xlsx", "xlsm", "xlsb", "xltx", "xltm", "xls", "xlt", "xml", "xlam", "xla", "xlw", "xlr", "ods")

Dim checkopenable As Boolean
checkopenable = False

Dim item As Variant

For Each item In filetype
    If Cells(rownum, Range("Col_FileType").Column) = item Then
        checkopenable = True
        Exit For
    End If
Next item
   
FileOpenable = checkopenable

'Dim tempstr, item As Variant
'tempstr = ""
'For Each item In filetype
'tempstr = tempstr & item & vbNewLine
'Next item
'MsgBox tempstr

End Function


Sub clear_control_input()

Dim start_row, end_row, i As Integer
    start_row = Range("Col_FullPath").Row + 1
    end_row = Range("Col_FullPath").End(xlDown).Row

Dim startcell, endcell As String
    startcell = Cells(start_row, Range("Col_index").Column).Address
    'endcell = Cells(end_row, Range("Col_Exp_Time").Column).Address
    endcell = Cells(1048000, Range("Col_Remarks").Column).Address
    
Range(startcell & ":" & endcell).ClearContents

End Sub


'Template Code
Sub CheckFolderExists()

Dim strFolderName As String
Dim strFolderExists As String

    strFolderName = "D:\Temp\Tool\CSV Gen\working\"
    strFolderExists = Dir(strFolderName, vbDirectory)

    If strFolderExists = "" Then
        MsgBox "The selected folder doesn't exist"
    Else
        MsgBox "The selected folder exists"
    End If

End Sub


'================================================================================================

Sub deletesheet()

Application.DisplayAlerts = False
Sheets(2).Delete
Application.DisplayAlerts = True

End Sub
