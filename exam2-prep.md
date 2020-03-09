# Exam II Preparation

Things you must be able to do:

    1. Display the results of a macro in a worksheet cell, message box or the immediate window

           
	    'Worksheet Cell
        '
        Range("A1")="Hello World"
        Sheets("Sheet2").Range("A5")="Hello World"
        '
	    '
        'Message Box
	    '
        MsgBox "Hello World"
	    '
        'Immediate Window
	    '
        Debug.Print "Hello World"
        '
    2. Find the last row
	    '
		lRow = Cells.Find(What:="*", _
                After:=Range("A1"), _
                LookAt:=xlPart, _
                LookIn:=xlFormulas, _
                SearchOrder:=xlByRows, _
                SearchDirection:=xlPrevious, _
                MatchCase:=False).Row
        '
        lRow = Cells(Rows.Count, 1).End(xlUp).Row
        '
        '
    3. Add a worksheet in a specific place with a specific name
	    '
        Dim dsp As Worksheet
        Set dsp = wb.Worksheets.Add(Before:=ip)
        dsp.Name = "DeathStarParts"
		'
		'
    4. Declare and use variables
	    '
        Dim wb As Workbook
        Set wb = ActiveWorkbook
        '
        Dim ip As Worksheet
        Set ip = wb.Sheets("inventory_parts")
        '
        Dim colors As Worksheet
        Set colors = wb.Sheets("colors")
        '
        Dim dscolors As Range
        Dim allcolors As Range
        '
		'
    5. Use looping statements such as do until, do while, and if
	    '
		' do until
		Do Until i > 6
        Cells(i, 1).Value = 20
        i = i + 1
		Loop
		'
		' do while
		'
        Do While i < 5
            i = i + 1
            msgbox "The value of i is : " & i
        Loop
		'
		' if and for
		'
        For Each c In allcolors
            For Each cc In dscolors
                If c.Value = cc.Value Then
                    c.Select
                    Set newsheetname = ActiveCell.Offset(0, 1)
                    Set newcolor = wb.Worksheets.Add(Before:=ip)
                    newcolor.Name = newsheetname.Value
                    Sheets("DeathStarParts").Select
                    Selection.AutoFilter
                    ActiveSheet.Range("$A$1:$L$592").AutoFilter Field:=6, Criteria1:=c.Value
                    Range("A1:L592").Select
                    Selection.Copy
                    Sheets(newsheetname.Value).Select
                    ActiveSheet.Paste
                    colors.Select
                End If
            Next cc
        Next c

    6. Sort and filter data from within a macro
        '
        Sheets("colors").Select
        Selection.AutoFilter
        ActiveWorkbook.Worksheets("colors").AutoFilter.Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("colors").AutoFilter.Sort.SortFields.Add2 Key:= _
            Range("B1:B136"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption _
            :=xlSortNormal
        With ActiveWorkbook.Worksheets("colors").AutoFilter.Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
		'
		'
    7. Be able to color, move, copy, or paste data based on given crtiteria or location
    8. All macros written for this exam must be able to be copied and pasted and used in any workbook.
