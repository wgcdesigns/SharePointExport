Sub TNGSharePointExport()

    '
    ' TNGSharePointExport
    ' 1/18/2018 NS
    ' Added urlAddress variable for easily changing the target url
    ' 1/17/2018 NS
    ' Script for exporting columns with large datasets in share point data to excel spreadsheet
    ' Currently looks for the keyword, "Progress Notes" in a specific group of cells
    '

    Dim i As Integer
    i = 9300
    Dim rc As Integer
    rc = 1
    Dim sheetrangebase As String
    sheetrangebase = "B"
    Dim sheetrangebasetitle As String
    sheetrangebasetitle = "A"
    Dim sheetrange As String
    sheetrange = ""
    Dim sheetrangetitle As String
    sheetrangetitle = ""
    Dim urlAddress As String
    urlAddress = "http://abc.com/default.aspx?ID="

    Do While i < 9311

    Sheets("Sheet1").Select
    Sheet1.Range("A1:A1000") = "" ' erase previous data
    Sheet1.Range("A1").Select

        With Sheet1.QueryTables.Add(Connection:=
            "URL;" & urlAddress & i, Destination:=Sheet1.Range("A1"))
            .Name = "default.aspx?ID=" & i
            .FieldNames = True
            .PreserveFormatting = True
            .BackgroundQuery = False
            .SaveData = True
            .AdjustColumnWidth = True
            .WebPreFormattedTextToColumns = True
            .WebConsecutiveDelimitersAsOne = True
            .Refresh "BackgroundQuery:=True"

            '.SavePassword = False
            '.RefreshOnFileOpen = False
            '.RowNumbers = False
            '.FillAdjacentFormulas = False
            '.WebSingleBlockTextImport = False
            '.WebDisableDateRecognition = False
            '.WebDisableRedirections = False
            '.RefreshStyle = xlInsertDeleteCells
            '.RefreshStyle = 0

            .RefreshPeriod = 0
            .WebSelectionType = xlEntirePage
            .WebFormatting = xlWebFormattingNone

            ' CREATE BASE STRINGS FOR PASTING INTO SHEET 2 --------
            sheetrange = sheetrangebase & rc
            sheetrangetitle = sheetrangebasetitle & rc

        End With

        ' WAIT UNTIL QUERY COMPLETE TO MOVE FORWARD ----------
        Dim StartTime As Single
        StartTime = Timer
        Do While Timer - StartTime <= 15
            'do thing
        Loop

        ' BREAKPOINT GOES HERE!!  TIMER ABOVE CAN BE CHANGED TO 5 INSTEAD TO INCREASE SPEED ALSO

        ' SELECT VALUES IN SHEET1 TO SEARCH FOR PROGRESS NOTES ---------------
        Sheets("Sheet1").Select
        Range("A55").Select

        ' IF VALUE IN CELL DOESNT EQUAL --- PROGRESS NOTES --- THEN FIND ALT CELL
        If InStr(1, Sheet1.Range("A55").Value, "Progress Notes") > 0 Then
            iFoundCell = "A55"
        ElseIf InStr(1, Sheet1.Range("A56").Value, "Progress Notes") > 0 Then
            Range("A56").Select
            iFoundCell = "A56"
        ElseIf InStr(1, Sheet1.Range("A57").Value, "Progress Notes") > 0 Then
            Range("A57").Select
            iFoundCell = "A57"
        ElseIf InStr(1, Sheet1.Range("A58").Value, "Progress Notes") > 0 Then
            Range("A58").Select
            iFoundCell = "A58"
        Else
            iFoundCell = "A59"
        End If

        Range(iFoundCell).Select
        Selection.Copy

        ' COPY PROGRESS NOTES IN SHEET 1 AND PASTE
        Sheets("Sheet2").Select
        Range(sheetrange).Select
        ActiveSheet.Paste
        ' PASTE PR NUMBER INTO A COLUMN -----------------
        Range(sheetrangetitle).Select
        Sheet2.Range(sheetrangetitle).Value = i
        '-----------------------------------
        i = i + 1
        rc = rc + 1
    Loop
   
End Sub
##
