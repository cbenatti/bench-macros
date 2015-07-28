Attribute VB_Name = "Stripe"
Option Explicit

Private Sub Stripe_Balance_History()
'Stripe Balance Holding
'Run this macro when you have a stripe balance_history report csv open
'This is the case where Stripe has a balance and is being treated as a bank account
'This macro transforms the stripe gross charges/transfers/refunds/adjustments into
'appropriate Bench-ready format with name and ledger, and sums fee column as well

'Variables
Dim TypeColumn As Long
Dim AmountColumn As Long
Dim FeeColumn As Long
Dim CreatedColumn As Long
Dim FinalRow As Long
Dim FinalColumn As String
Dim CurrentRow As Long
Dim TypeRange As Range
Dim AmountRange As Range
Dim CreatedRange As Range
Dim FeeSums As Variant
Dim StripeWorkbook As Workbook
Dim CurrentWorkbook As Workbook
Dim CurrentWorkbookPath As String
Dim CurrentWorkbookName As String
Dim StripeType As String
Dim RowYear As Long
Dim RowMonth As Long
Dim MonthsIter As Long
Dim MonthString As String
Dim i As Long

'Run the Error handler "ErrHandler" when an error occurs
On Error GoTo ErrHandler

'Turn off screen updating and display alerts
Application.ScreenUpdating = False
Application.DisplayAlerts = False
DoEvents

'Initialize final row/column variables
FinalRow = LastRow(ActiveSheet)
FinalColumn = ConvertToLetter(LastCol(ActiveSheet))

'Define columns
TypeColumn = FindColNumber(ActiveSheet, "Type", "A", FinalColumn)
AmountColumn = FindColNumber(ActiveSheet, "Amount", "A", FinalColumn)
FeeColumn = FindColNumber(ActiveSheet, "Fee", "A", FinalColumn)
CreatedColumn = FindColNumber(ActiveSheet, "Created (UTC)", "A", FinalColumn)
If CreatedColumn <> 0 Then
    'Do nothing, created column is found!
Else
    'No created date, search for "Date"
    CreatedColumn = FindColNumber(ActiveSheet, "Date", "A", FinalColumn)
End If

'Define ranges
Set TypeRange = ActiveSheet.Range(Cells(2, TypeColumn), Cells(FinalRow, TypeColumn))
Set AmountRange = ActiveSheet.Range(Cells(2, AmountColumn), Cells(FinalRow, AmountColumn))
Set CreatedRange = ActiveSheet.Range(Cells(2, CreatedColumn), Cells(FinalRow, CreatedColumn))

'Sort data by Created date so can calculated monthly fee sums
Range(Cells(1, 1), Cells(FinalRow, FinalColumn)).Sort Key1:=CreatedRange, Order1:=xlAscending, Header:=xlYes

'Calculate fee sum
'FeeSums has three rows, and as many columns as unique months in the report
'First row is the year
'Second row is the month
'Third row is the sum (cumulative)
'Would have made it columns instead of rows, but VBA can only redim last dimension

'Initialize FeeSums with first row of data
ReDim FeeSums(2, 0) As Double
FeeSums(0, 0) = Year(Cells(2, CreatedColumn))                         'Year
FeeSums(1, 0) = Month(Cells(2, CreatedColumn))                      'Month
FeeSums(2, 0) = -1 * Cells(2, FeeColumn).Value                       'FeeSum
CurrentRow = 3
MonthsIter = 0 'Total Months in report is MonthsIter+1

Do While CurrentRow <= FinalRow
    RowYear = Year(Cells(CurrentRow, CreatedColumn))
    RowMonth = Month(Cells(CurrentRow, CreatedColumn))
    If RowYear = FeeSums(0, MonthsIter) Then
        'Same year, continue looking
        If RowMonth = FeeSums(1, MonthsIter) Then
            'Same month, so it's a match!
            FeeSums(2, MonthsIter) = FeeSums(2, MonthsIter) + Cells(CurrentRow, FeeColumn).Value 'Maybe should subtract?
        Else
            'Different month, so a new row
            MonthsIter = MonthsIter + 1
            ReDim Preserve FeeSums(2, MonthsIter)
            FeeSums(0, MonthsIter) = RowYear
            FeeSums(1, MonthsIter) = RowMonth
            FeeSums(2, MonthsIter) = Cells(CurrentRow, FeeColumn).Value 'Should this be negative?
        End If
    Else
        'Different year, so a new row
        MonthsIter = MonthsIter + 1
        ReDim Preserve FeeSums(2, MonthsIter)
        FeeSums(0, MonthsIter) = RowYear
        FeeSums(1, MonthsIter) = RowMonth
        FeeSums(2, MonthsIter) = Cells(CurrentRow, FeeColumn).Value 'Should this be negative?
    End If
    CurrentRow = CurrentRow + 1
'FeeSum = WorksheetFunction.Sum(ActiveSheet.Range(Cells(2, FeeColumn), Cells(FinalRow, FeeColumn)))
Loop
 
CurrentRow = 2

'Get current workbook path and name
Set CurrentWorkbook = ActiveWorkbook
CurrentWorkbookPath = ActiveWorkbook.Path
CurrentWorkbookName = ActiveWorkbook.Name

'Create new workbook
Set StripeWorkbook = Workbooks.Add
StripeWorkbook.Activate

'Column headers
Range("A1").Value = "Company"
Range("B1").Value = "Date of transaction"
Range("C1").Value = "Amount"
Range("D1").Value = "Ledger"

'Copy columns into new workbook
TypeRange.Copy
Range("A2").PasteSpecial xlPasteValues
CreatedRange.Copy
Range("B2").PasteSpecial xlPasteValues
AmountRange.Copy
Range("C2").PasteSpecial xlPasteValues

'Loop through rows, fixing transaction name and ledger
Do While CurrentRow <= FinalRow
    StripeType = Range("A" + CStr(CurrentRow)).Value
    If StripeType = "charge" Then
        Range("A" + CStr(CurrentRow)).Value = "Stripe Charge"
        Range("D" + CStr(CurrentRow)).Value = "Sales Revenue"
    ElseIf StripeType = "refund" Or StripeType = "adjustment" Then
        Range("A" + CStr(CurrentRow)).Value = "Stripe Refund"
        Range("D" + CStr(CurrentRow)).Value = "Sales Returns"
    ElseIf StripeType = "transfer" Then
        Range("A" + CStr(CurrentRow)).Value = "Stripe Transfer"
        Range("D" + CStr(CurrentRow)).Value = "Money in Transit (outstanding)"
    End If
    CurrentRow = CurrentRow + 1
Loop

'Fee inserted into new workbook
For i = 0 To MonthsIter
    MonthString = Format(DateSerial(FeeSums(0, i), FeeSums(1, i), 1), "mmmm")
    Range("A" + CStr(FinalRow + 1 + i)).Value = "Monthly Stripe Adjustment - " + MonthString
    Range("B" + CStr(FinalRow + 1 + i)).Value = dhLastDayInMonth(DateSerial(FeeSums(0, i), FeeSums(1, i), 1))
    Range("C" + CStr(FinalRow + 1 + i)).Value = FeeSums(2, i)
    Range("D" + CStr(FinalRow + 1 + i)).Value = "Merchant Fees Expense"
Next i

'Autofit and formatcolumns
Columns("A").Select
Columns("A").AutoFit
Columns("B").Select
Columns("B").AutoFit
Columns("B").NumberFormat = "yyyy-mm-dd;@"
Columns("C").Select
'Columns("C").AutoFit 'don't autofit amount, gets too small!
Columns("C").NumberFormat = "0.00;@"
Columns("D").Select
Columns("D").AutoFit

'Last Selection
Range("A1").Select

'Turn on screen updating and display alerts
Application.ScreenUpdating = True
Application.DisplayAlerts = True
DoEvents

'Save workbook as an xlsx instead of csv
On Error Resume Next
ActiveWorkbook.SaveAs CurrentWorkbookPath + ":" + Replace(CurrentWorkbookName, ".csv", ".xlsx"), FileFormat:=52
        
'Exit the macro so that the error handler is not executed
Exit Sub

ErrHandler:
    'If an error occurs, display a message and end the macro
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
               vbExclamation, "Stripe Balance History"

End Sub

Private Function ColumnName_gibbo(cellRange As Range) As String
    ColumnName_gibbo = Split(cellRange.Address(1, 0), "$")(0)
End Function

Private Function FindRowNumber(CurrentWorksheet As Worksheet, findText As String, rowStart As Long, rowEnd As Long) As Long
    'Finds text in column A
    Dim FindRow As Range
            
    Set FindRow = CurrentWorksheet.Range("A" + CStr(rowStart) + ":A" + CStr(rowEnd)).Find(What:=findText, LookIn:=xlValues, MatchCase:=True)
    
    If FindRow Is Nothing Then
        FindRowNumber = 0
    Else
        FindRowNumber = FindRow.Row
    End If
    
End Function

Private Function FindColNumber(CurrentWorksheet As Worksheet, findText As String, colStart As String, colEnd As String) As Long
    'Finds text in row 1
    Dim FindCol As Range
    Dim FinalCol As Long
    Dim CurrentCol As Long
    
    Set FindCol = CurrentWorksheet.Range(colStart + "1:" + colEnd + "1").Find(What:=findText, LookIn:=xlValues, MatchCase:=True)
    
    If FindCol Is Nothing Then
        FindColNumber = 0
    Else
        FindColNumber = FindCol.Column
    End If
    
     'Make sure to find only exact matches, and not where a word could be found within a cell with other words
    If (findText = "Total" Or findText = "TOTAL") And FindColNumber <> 0 Then
        If UCase(Range(ConvertToLetter(CInt(FindColNumber)) + "1").Text) = "TOTAL" Then
            'do nothing
        Else
            FinalCol = LastCol(CurrentWorksheet)
            CurrentCol = FindColNumber + 1
            Do While CurrentCol <= FinalCol
                Set FindCol = CurrentWorksheet.Range(ConvertToLetter(CInt(CurrentCol)) + "1:" + ConvertToLetter(CInt(FinalCol)) + "1").Find(What:=findText, LookIn:=xlValues, MatchCase:=True)
                If FindCol Is Nothing Then
                    FindColNumber = 0
                    Exit Function
                Else
                    FindColNumber = FindCol.Column
                End If
                If UCase(Range(ConvertToLetter(CInt(FindColNumber)) + "1").Text) = "TOTAL" Then
                    Exit Function
                Else
                    'do nothing
                End If
                CurrentCol = FindColNumber + 1
            Loop
        End If
    Else
        'do nothing
    End If
    If (findText = "Date" Or findText = "DATE") And FindColNumber <> 0 Then
        If UCase(Range(ConvertToLetter(CInt(FindColNumber)) + "1").Text) = "DATE" Then
            'do nothing
        Else
            FinalCol = LastCol(CurrentWorksheet)
            CurrentCol = FindColNumber + 1
            Do While CurrentCol <= FinalCol
                Set FindCol = CurrentWorksheet.Range(ConvertToLetter(CInt(CurrentCol)) + "1:" + ConvertToLetter(CInt(FinalCol)) + "1").Find(What:=findText, LookIn:=xlValues, MatchCase:=True)
                If FindCol Is Nothing Then
                    FindColNumber = 0
                    Exit Function
                Else
                    FindColNumber = FindCol.Column
                End If
                If UCase(Range(ConvertToLetter(CInt(FindColNumber)) + "1").Text) = "DATE" Then
                    Exit Function
                Else
                    'do nothing
                End If
                CurrentCol = FindColNumber + 1
            Loop
        End If
    Else
        'do nothing
    End If
    
End Function

Private Function LastRow(WS As Object) As Long

        Dim rLastCell As Object
        On Error GoTo ErrHan
        Set rLastCell = WS.Cells.Find("*", WS.Cells(1, 1), , , xlByRows, _
                                      xlPrevious)
        LastRow = rLastCell.Row

ErrExit:
        Exit Function

ErrHan:
        'MsgBox "Error " & Err.Number & ": " & Err.Description, _
        '       vbExclamation, "LastRow()"
        Resume ErrExit

End Function

Private Function LastCol(WS As Object) As Long

        Dim rLastCell As Object
        On Error GoTo ErrHan
        Set rLastCell = WS.Cells.Find("*", WS.Cells(1, 1), , , xlByColumns, _
                                      xlPrevious)
        LastCol = rLastCell.Column

ErrExit:
        Exit Function

ErrHan:
        'MsgBox "Error " & Err.Number & ": " & Err.Description, _
        '       vbExclamation, "LastRow()"
        Resume ErrExit

End Function

Private Function ConvertToLetter(iCol As Integer) As String
   Dim iAlpha As Integer
   Dim iRemainder As Integer
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = Chr(iAlpha + 64)
   End If
   If iRemainder > 0 Then
      ConvertToLetter = ConvertToLetter & Chr(iRemainder + 64)
   End If
End Function

Private Function dhFirstDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the first day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhFirstDayInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate), 1)
End Function

Private Function dhLastDayInMonth(Optional dtmDate As Date = 0) As Date
    ' Return the last day in the specified month.
    If dtmDate = 0 Then
        ' Did the caller pass in a date? If not, use
        ' the current date.
        dtmDate = Date
    End If
    dhLastDayInMonth = DateSerial(Year(dtmDate), _
     Month(dtmDate) + 1, 0)
End Function

Private Sub CompareStripeAndBenchSearchResults()
'This sub reads in all the transactions from a Bench "Search-Results" file
'as well as all the transactions from a Stripe "Stripe Transfers" file
'Both files must be open at the time when this macro is run
'A new file is created, comparing the two transactions and outlining any
'transactions from the Bench file that do not exist in the Stripe file

Dim BenchWorkbook As Workbook
Dim StripeWorkbook As Workbook
Dim ReportWorkbook As Workbook
Dim FinalFileName As String
Dim i As Long
Dim MonthsIter As Long

'Run the Error handler "ErrHandler" when an error occurs
On Error GoTo ErrHandler

'Define Workbooks
For i = 1 To Workbooks.Count
    'Find BenchWorkbook (contains "Search-Results" in name)
    If InStr(1, Workbooks(i).Name, "Search-Results", vbTextCompare) > 0 Then
        Set BenchWorkbook = Workbooks(i)
    'Find StripeWorkbook (contains "Stripe Transfers" in name)
    ElseIf InStr(1, Workbooks(i).Name, "Stripe Transfers", vbTextCompare) > 0 Then
        Set StripeWorkbook = Workbooks(i)
    End If
Next i

'Check that both BenchWorkbook and StripeWorkbook have been set
If BenchWorkbook Is Nothing Then
    MsgBox "Need to have ""Search-Results"" file open"
    Exit Sub
ElseIf StripeWorkbook Is Nothing Then
    'MsgBox "Need to have ""Stripe Transfers"" file open"
    'Exit Sub
    Application.Run "CombineFolderCSVs.CombineCSVsInUserDefinedFolder"
    Set StripeWorkbook = ActiveWorkbook
End If

'Create new workbook
Set ReportWorkbook = Workbooks.Add
ReportWorkbook.Activate

'Turn off screen updating and display alerts
Application.ScreenUpdating = False
Application.DisplayAlerts = False
DoEvents

'Copy over Stripe data - return number of months (use as input when copying in bench data)
MonthsIter = CopyStripeData(StripeWorkbook, ReportWorkbook)

'Copy over Bench data
CopyBenchData BenchWorkbook, ReportWorkbook, MonthsIter

'Turn on screen updating and display alerts
Application.ScreenUpdating = True
Application.DisplayAlerts = True
DoEvents

'Save report workbook
FinalFileName = BenchWorkbook.Path + ":" + "StripeBenchCompare.xlsx"
On Error Resume Next
ActiveWorkbook.SaveAs FinalFileName, FileFormat:=52

'Exit sub so error handler is not called
Exit Sub

ErrHandler:
    'If an error occurs, display a message and end the macro
    MsgBox "Error " & Err.Number & ": " & Err.Description, _
               vbExclamation, "Stripe Balance History"

End Sub

Private Function CopyStripeData(FromWorkbook As Workbook, ToWorkbook As Workbook) As Long
'Copies data from "Stripe Transfers" file into report workbook]

'Variables
Dim DescriptionColumn As Long
Dim AmountColumn As Long
Dim FeeColumn As Long
Dim CreatedColumn As Long
Dim GrossColumn As Long
Dim FinalRow As Long
Dim FinalColumn As String
Dim CurrentRow As Long
Dim DescriptionRange As Range
Dim AmountRange As Range
Dim CreatedRange As Range
Dim FeeRange As Range
Dim FeeSums As Variant
Dim RowYear As Long
Dim RowMonth As Long
Dim MonthsIter As Long
Dim MonthString As String
Dim i As Long
Dim GrossSum As Double

On Error GoTo 0

'Initialize final row/column variables
FinalRow = LastRow(FromWorkbook.Sheets(1))
FinalColumn = ConvertToLetter(LastCol(FromWorkbook.Sheets(1)))

'Define columns
FromWorkbook.Activate
DescriptionColumn = FindColNumber(FromWorkbook.Sheets(1), "Description", "A", FinalColumn)
AmountColumn = FindColNumber(FromWorkbook.Sheets(1), "Total", "A", FinalColumn)
FeeColumn = FindColNumber(FromWorkbook.Sheets(1), "Total Fees", "A", FinalColumn)
GrossColumn = FindColNumber(FromWorkbook.Sheets(1), "Total Gross", "A", FinalColumn)
CreatedColumn = FindColNumber(FromWorkbook.Sheets(1), "Date", "A", FinalColumn)
If CreatedColumn <> 0 Then
    'Do nothing, created column is found!
Else
    'No created date, search for "Created"
    CreatedColumn = FindColNumber(ActiveSheet, "Created (UTC)", "A", FinalColumn)
End If
If AmountColumn <> 0 Then
    'Do nothing, amount column is found!
Else
    'No created date, search for "Amount"
    AmountColumn = FindColNumber(ActiveSheet, "Tax", "A", FinalColumn) 'Amount... or tax? is column header messed up in sophie's file?
End If
If FeeColumn <> 0 Then
    'Do nothing, amount column is found!
Else
    'No created date, search for "Fee"
    FeeColumn = FindColNumber(ActiveSheet, "Fee", "A", FinalColumn)
End If
If GrossColumn <> 0 Then
    'Calculate sum, gross total column is found!
    GrossSum = Application.Sum(ActiveSheet.Range(ConvertToLetter(CInt(GrossColumn)) + "2:" + ConvertToLetter(CInt(GrossColumn)) + CStr(FinalRow)))
Else
    'No gross total column found - set sum to zero
    GrossSum = 0
End If


'Define ranges
FromWorkbook.Activate
Set DescriptionRange = FromWorkbook.Sheets(1).Range(Cells(2, DescriptionColumn), Cells(FinalRow, DescriptionColumn))
Set AmountRange = FromWorkbook.Sheets(1).Range(Cells(2, AmountColumn), Cells(FinalRow, AmountColumn))
Set CreatedRange = FromWorkbook.Sheets(1).Range(Cells(2, CreatedColumn), Cells(FinalRow, CreatedColumn))
Set FeeRange = FromWorkbook.Sheets(1).Range(Cells(2, FeeColumn), Cells(FinalRow, FeeColumn))

'Column headers in ToWorkbook
ToWorkbook.Activate
ToWorkbook.Sheets(1).Range("A1").Value = "Stripe Description"
ToWorkbook.Sheets(1).Range("B1").Value = "Date"
ToWorkbook.Sheets(1).Range("C1").Value = "Amount"
'Temporary fee column
ToWorkbook.Sheets(1).Range("D1").Value = "Fees"
'GrossSum
ToWorkbook.Sheets(1).Range("N1").Value = "Total Gross Sum"

'Copy columns into ToWorkbook
DescriptionRange.Copy
ToWorkbook.Sheets(1).Range("A2").PasteSpecial xlPasteValues
CreatedRange.Copy
ToWorkbook.Sheets(1).Range("B2").PasteSpecial xlPasteValues
AmountRange.Copy
ToWorkbook.Sheets(1).Range("C2").PasteSpecial xlPasteValues
'Temporarily copy over fees column
FeeRange.Copy
ToWorkbook.Sheets(1).Range("D2").PasteSpecial xlPasteValues

'GrossSum from stripe workbook
ToWorkbook.Sheets(1).Range("N2").Value = GrossSum

'Redefine final row and column, and created column and range
FinalRow = LastRow(ToWorkbook.Sheets(1))
FinalColumn = "D"
CreatedColumn = 2
Set CreatedRange = ToWorkbook.Sheets(1).Range(Cells(2, CreatedColumn), Cells(FinalRow, CreatedColumn))
FeeColumn = 4

'Sort data by Created date so can calculated monthly fee sums
ToWorkbook.Sheets(1).Range(Cells(1, 1), Cells(FinalRow, FinalColumn)).Sort Key1:=CreatedRange, Order1:=xlAscending, Header:=xlYes

'Calculate fee sum
'FeeSums has three rows, and as many columns as unique months in the report
'First row is the year
'Second row is the month
'Third row is the sum (cumulative)
'Would have made it columns instead of rows, but VBA can only redim last dimension

'Initialize FeeSums with first row of data
ReDim FeeSums(2, 0) As Double
FeeSums(0, 0) = Year(Cells(2, CreatedColumn))                         'Year
FeeSums(1, 0) = Month(Cells(2, CreatedColumn))                      'Month
FeeSums(2, 0) = Cells(2, FeeColumn).Value                               'FeeSum
CurrentRow = 3
MonthsIter = 0 'Total Months in report is MonthsIter+1

Do While CurrentRow <= FinalRow
    RowYear = Year(Cells(CurrentRow, CreatedColumn))
    RowMonth = Month(Cells(CurrentRow, CreatedColumn))
    If RowYear = FeeSums(0, MonthsIter) Then
        'Same year, continue looking
        If RowMonth = FeeSums(1, MonthsIter) Then
            'Same month, so it's a match!
            FeeSums(2, MonthsIter) = FeeSums(2, MonthsIter) + Cells(CurrentRow, FeeColumn).Value
        Else
            'Different month, so a new row
            MonthsIter = MonthsIter + 1
            ReDim Preserve FeeSums(2, MonthsIter)
            FeeSums(0, MonthsIter) = RowYear
            FeeSums(1, MonthsIter) = RowMonth
            FeeSums(2, MonthsIter) = Cells(CurrentRow, FeeColumn).Value
        End If
    Else
        'Different year, so a new row
        MonthsIter = MonthsIter + 1
        ReDim Preserve FeeSums(2, MonthsIter)
        FeeSums(0, MonthsIter) = RowYear
        FeeSums(1, MonthsIter) = RowMonth
        FeeSums(2, MonthsIter) = Cells(CurrentRow, FeeColumn).Value
    End If
    CurrentRow = CurrentRow + 1
Loop
 
CurrentRow = 2

'Fee inserted into new workbook
For i = 0 To MonthsIter
    MonthString = Format(DateSerial(FeeSums(0, i), FeeSums(1, i), 1), "mmmm")
    Range("A" + CStr(FinalRow + 1 + i)).Value = "Monthly Stripe Adjustment - " + MonthString
    Range("B" + CStr(FinalRow + 1 + i)).Value = dhLastDayInMonth(DateSerial(FeeSums(0, i), FeeSums(1, i), 1))
    Range("C" + CStr(FinalRow + 1 + i)).Value = FeeSums(2, i)
Next i

'Autofit and formatcolumns
Columns("A").Select
Columns("A").AutoFit
Columns("B").Select
Columns("B").NumberFormat = "yyyy-mm-dd;@"
Columns("B").AutoFit
Columns("C").Select
'Columns("C").AutoFit 'don't autofit amount, gets too small!
Columns("C").NumberFormat = "0.00;@"
Columns("N").Select
Columns("N").AutoFit
Columns("N").NumberFormat = "0.00;@"

'Delete Fees Column
'Columns("D").Delete

'Redefine final row and column, and created column and range
FinalRow = LastRow(ToWorkbook.Sheets(1))
FinalColumn = "D"
CreatedColumn = 2
Set CreatedRange = ToWorkbook.Sheets(1).Range(Cells(2, CreatedColumn), Cells(FinalRow, CreatedColumn))

'Sort data by Created date
ToWorkbook.Sheets(1).Range(Cells(1, 1), Cells(FinalRow, FinalColumn)).Sort Key1:=CreatedRange, Order1:=xlAscending, Header:=xlYes

'Last Selection
Range("A1").Select

CopyStripeData = MonthsIter
End Function

Private Sub CopyBenchData(FromWorkbook As Workbook, ToWorkbook As Workbook, MonthsIter As Long)
'Copies Bench data from "Search-Results" file, and compares with previously
'copied in data from stripe workbook -- unmatched transactions on the right

Dim BenchSearchData As Variant
Dim TotalBenchRows As Long
Dim FinalRow As Long 'Final row in ToWorkbook
Dim RowString As String
Dim RowString2 As String
Dim RowString3 As String
Dim FoundRow As Long
Dim SearchRange As Range
Dim SearchStartRow As Long
Dim UnmatchedRow As Long
Dim i As Long
Dim j As Long
Dim BenchMonthStartRow As Long
Dim BenchMonthEndRow As Long
Dim CurrentMonth As Long
Dim CurrentMonthRow As Long
Dim FindString As String
Dim DateTolerance As Long
Dim FoundMatch As Boolean
Dim HighlightingState As Boolean 'True=highlight stripe and bench only for doubles/triples 'False=highlight stripe column at end
HighlightingState = False
Dim DateLowerLimit As Date
Dim DateUpperLimit As Date

On Error GoTo 0

'Set Date tolerance (in days -- what constitues a match)
DateTolerance = 5

'Get total rows in FromWorkbook
FromWorkbook.Activate
TotalBenchRows = LastRow(FromWorkbook.Sheets(1))
    
'Work with FromWorkbook
With FromWorkbook.Sheets(1)

    'Use a better date format for copying
    If TypeName(FromWorkbook.Sheets(1).Range("A4").Value) = "String" Then
        'Need to do Text to Columns First to make it a date
        .Range("A4:A" + CStr(TotalBenchRows)).Select
        TTC_DMY
    Else
        'Already a date, good
    End If
    .Range("A4:A" + CStr(TotalBenchRows)).NumberFormat = "yyyy-mm-dd;@"
    .Range("A1").Select
    
    'Populate BenchSearchData
    ReDim BenchSearchData(2, TotalBenchRows - 4)
    For i = 0 To TotalBenchRows - 4
        RowString = CStr(i + 4)
        BenchSearchData(0, i) = .Range("A" + RowString).Value 'Date
        BenchSearchData(1, i) = .Range("B" + RowString).Value 'Description
        If .Range("D" + RowString).Value <> "Sales Revenue" Then
            BenchSearchData(2, i) = .Range("E" + RowString).Value 'Amount
        Else
            BenchSearchData(2, i) = -1 * .Range("E" + RowString).Value 'Amount
        End If
    Next i
End With

'Get total rows in ToWorkbook
ToWorkbook.Activate
FinalRow = LastRow(ToWorkbook.Sheets(1))

'Work with ToWorkbook
With ToWorkbook.Sheets(1)
    'Column headers in ToWorkbook
    .Range("E1").Value = "Bench Description"
    .Range("F1").Value = "Date"
    .Range("G1").Value = "Amount"
    .Range("I1").Value = "Unmatched from Bench"
    .Range("J1").Value = "Date"
    .Range("K1").Value = "Amount"
    .Range("M1").Value = "Amount Sum"
    
    'Initialize row variables
    SearchStartRow = 2
    UnmatchedRow = 2
    
    'Sort Backwards by Stripe Date (for searching, want to search older first) and determine upper/lower date limits
    .Range(Cells(1, 1), Cells(FinalRow, 4)).Sort Key1:=.Range(Cells(2, 2), Cells(FinalRow, 2)), Order1:=xlDescending, Header:=xlYes
    DateUpperLimit = .Range("B" + CStr(FinalRow)).Value + 5
    DateLowerLimit = .Range("B2").Value - 5
    
    'Try to find transaction within ReportWorkbook (ToWorkbook)
    For i = 0 To TotalBenchRows - 4
        'Check that Bench date (ie one searching for, is within range: SmallestStripeDate-5<=BenchDate<=LargestStripeDate+5
        If DateLowerLimit <= BenchSearchData(2, i) <= DateUpperLimit Then
            
            'Define SearchRange
            Set SearchRange = .Range(Cells(SearchStartRow, 3), Cells(FinalRow, 3))
            
            On Error Resume Next
            FoundRow = 0
            FoundRow = Application.WorksheetFunction.Match(BenchSearchData(2, i), SearchRange, 0)
            On Error GoTo 0
                         
            If FoundRow > 0 Then
                'Match was found - but relative to search range, so need to get actual found row value
                FoundRow = FoundRow + SearchStartRow - 1
                
                'Matched a transaction amount -- still need to do error handling here
                'Check that date is within the date tolerance
                RowString = CStr(FoundRow)
                If Abs(DateDiff("d", BenchSearchData(0, i), .Range("B" + RowString).Value)) <= DateTolerance Then
                    'Check that a match has not already been found
                    If IsEmpty(ToWorkbook.Sheets(1).Range("E" + RowString)) Then
                        .Range("E" + RowString).Value = BenchSearchData(1, i)
                        .Range("F" + RowString).Value = BenchSearchData(0, i)
                        .Range("G" + RowString).Value = BenchSearchData(2, i)
                        SearchStartRow = 2
                    Else
                        'A match has already been found for that amount and date! Continue searching
                        SearchStartRow = FoundRow + 1
                        i = i - 1 'Repeat for same BenchSearchData row
                    End If
                 Else
                    'Continue searching
                    SearchStartRow = FoundRow + 1
                    i = i - 1 'Repeat for same BenchSearchData row
                 End If
            Else
                FoundMatch = False
                'Did not match a single transaction, so look at doubles (and then triples possibly?)
                'For now, algorithm requires that doubles are next to each other
                For j = 2 To FinalRow - 1
                    RowString = CStr(j)
                    RowString2 = CStr(j + 1)
                    'Check that match has not already been found
                    If IsEmpty(.Range("E" + RowString)) And _
                       IsEmpty(.Range("E" + RowString2)) Then
                        'Possibility for a match
                        If Abs(BenchSearchData(2, i) - (.Range("C" + RowString).Value + .Range("C" + RowString2).Value)) < 0.01 Then
                            'It's a match!
                            'Check the dates
                            If Abs(DateDiff("d", BenchSearchData(0, i), .Range("B" + RowString).Value)) <= DateTolerance And _
                               Abs(DateDiff("d", BenchSearchData(0, i), .Range("B" + RowString2).Value)) <= DateTolerance Then
                                'The dates are close enough!
                                .Range("E" + RowString2).Value = BenchSearchData(1, i)
                                .Range("F" + RowString2).Value = BenchSearchData(0, i)
                                .Range("G" + RowString2).Value = BenchSearchData(2, i)
                                SearchStartRow = 2
                                
                                If HighlightingState Then
                                    'Highlight the sum in yellow, as well as the two values that add together
                                    .Range("G" + RowString2).Interior.Color = vbYellow
                                    .Range("C" + RowString).Interior.Color = vbYellow
                                    .Range("C" + RowString2).Interior.Color = vbYellow
                                End If
                                
                                'Put " centered into RowString row (this will end up being the second row when sorted forwards)
                                .Range("E" + RowString + ":G" + RowString).Value = """"
                                .Range("E" + RowString + ":G" + RowString).HorizontalAlignment = xlCenter
                                
                                'Set found match boolean state
                                FoundMatch = True
                                Exit For
                            Else
                                'Move on to the next possible match
                            End If
                        Else
                            'Move on to next possible match
                        End If
                    Else
                        'Move on to next possible match
                    End If
                Next j
                If FoundMatch Then
                    'Do Nothing
                Else
                    'Check for triples
                    'For now, algorithm requires that triples are next to each other
                    For j = 2 To FinalRow - 2
                        RowString = CStr(j)
                        RowString2 = CStr(j + 1)
                        RowString3 = CStr(j + 2)
                        'Check that match has not already been found
                        If IsEmpty(.Range("E" + RowString)) And _
                           IsEmpty(.Range("E" + RowString2)) And _
                           IsEmpty(.Range("E" + RowString3)) Then
                            'Possibility for a match
                            If BenchSearchData(2, i) = .Range("C" + RowString).Value + .Range("C" + RowString2).Value + .Range("C" + RowString3).Value Then
                                'It's a match!
                                'Check the dates
                                If Abs(DateDiff("d", BenchSearchData(0, i), .Range("B" + RowString).Value)) <= DateTolerance And _
                                   Abs(DateDiff("d", BenchSearchData(0, i), .Range("B" + RowString2).Value)) <= DateTolerance And _
                                   Abs(DateDiff("d", BenchSearchData(0, i), .Range("B" + RowString3).Value)) <= DateTolerance Then
                                    'The dates are close enough!
                                    .Range("E" + RowString3).Value = BenchSearchData(1, i)
                                    .Range("F" + RowString3).Value = BenchSearchData(0, i)
                                    .Range("G" + RowString3).Value = BenchSearchData(2, i)
                                    SearchStartRow = 2
                                    
                                    If HighlightingState Then
                                        'Highlight the sum in cyan, as well as the three values that add together
                                        .Range("G" + RowString3).Interior.Color = vbCyan
                                        .Range("C" + RowString).Interior.Color = vbCyan
                                        .Range("C" + RowString2).Interior.Color = vbCyan
                                        .Range("C" + RowString3).Interior.Color = vbCyan
                                    End If
                                    
                                    'Put " centered into RowString and RowString2 row (this will end up being the second/third rows when sorted forwards)
                                    .Range("E" + RowString + ":G" + RowString2).Value = """"
                                    .Range("E" + RowString + ":G" + RowString2).HorizontalAlignment = xlCenter
                                    
                                    'Set found match boolean state
                                    FoundMatch = True
                                    Exit For
                                Else
                                    'Move on to the next possible match
                                End If
                            Else
                                'Move on to next possible match
                            End If
                        Else
                            'Move on to next possible match
                        End If
                    Next j
                    If FoundMatch Then
                        'Do nothing
                    Else
                        'No matching transaction
                        RowString = CStr(UnmatchedRow)
                        .Range("I" + RowString).Value = BenchSearchData(1, i)
                        .Range("J" + RowString).Value = BenchSearchData(0, i)
                        .Range("K" + RowString).Value = BenchSearchData(2, i)
                        UnmatchedRow = UnmatchedRow + 1
                        SearchStartRow = 2
                    End If
                End If
            End If
        Else
            'No matching transaction
            RowString = CStr(UnmatchedRow)
            .Range("I" + RowString).Value = BenchSearchData(1, i)
            .Range("J" + RowString).Value = BenchSearchData(0, i)
            .Range("K" + RowString).Value = BenchSearchData(2, i)
            UnmatchedRow = UnmatchedRow + 1
            SearchStartRow = 2
        End If
    Next i
    
    'ColumnF must be formatted this way to search for month (the way I've written it at least)
    .Columns("F").NumberFormat = "yyyy-mm-dd;@"
    
    'Recalculate Monthly adjusted fees based on matched Bench dates
    CurrentMonthRow = 2
    For i = 0 To MonthsIter
        CurrentMonthRow = FindRowNumber(ToWorkbook.Sheets(1), "Monthly Stripe Adjustment", CurrentMonthRow, FinalRow)
        If CurrentMonthRow > 0 Then
            CurrentMonth = Month(.Range("B" + CStr(CurrentMonthRow)).Value)
            FindString = Format(CurrentMonth, "-00-")
            
            'Find first instance of current month in bench column
            If Range("F:F").Find(FindString, after:=.Range("F1")) Is Nothing Then
                Exit For
            Else
                BenchMonthStartRow = Range("F:F").Find(FindString, after:=.Range("F1")).Row
            End If
            
            'Find last instance of current month in bench column
            If Range("F:F").Find(FindString, after:=.Range("F1")) Is Nothing Then
                Exit For
            Else
                BenchMonthEndRow = .Range("F:F").Find(FindString, after:=.Range("F1"), searchdirection:=xlPrevious).Row
            End If
            
            'Sum fees within that range, and put into Monthly Stripe Adjustment cell
            .Range("C" + CStr(CurrentMonthRow)).Value = Application.Sum(.Range("D" + CStr(BenchMonthStartRow) + ":D" + CStr(BenchMonthEndRow)))
            
            'Check if a previously MATCHED transaction now doesn't match the new monthly stripe adjustment value
            If Not IsEmpty(.Range("E" + CStr(CurrentMonthRow))) Then
                'There was a previous match -- check if it's still a match
                If Abs(.Range("C" + CStr(CurrentMonthRow)).Value - .Range("G" + CStr(CurrentMonthRow)).Value) <= 0.01 Then
                    'It's still a match -- do nothing
                Else
                    'It's no longer a match -- move into unmatched columns
                    .Range("E" + CStr(CurrentMonthRow) + ":G" + CStr(CurrentMonthRow)).Cut
                    .Range("H" + CStr(.Cells(.Rows.Count, "H").End(xlUp).Row + 1)).PasteSpecial xlPasteValues
                    'Search if previously unmatched transaction now matches
                    SearchPreviouslyUnmatchedTransactions ToWorkbook.Sheets(1), CurrentMonthRow, DateTolerance
                End If
                
            Else
                'There was not a match previously
                'Search if a previously unmatched transaction from bench now matches the new monthly stripe adjustment value
                SearchPreviouslyUnmatchedTransactions ToWorkbook.Sheets(1), CurrentMonthRow, DateTolerance
            End If
            
            'Increment current month row to search for next monthly stripe adjustment
            CurrentMonthRow = CurrentMonthRow + 1
        Else
            'No additional Month Stripe Adjustments found
            Exit For
        End If
        
    Next i
       
    'Delete Fees column
    .Columns("D").Delete
    
    'Sort Forwards by Stripe Date (for searching, want to search older first)
    .Range(Cells(1, 1), Cells(FinalRow, 6)).Sort Key1:=.Range(Cells(2, 2), Cells(FinalRow, 2)), Order1:=xlAscending, Header:=xlYes
    .Range(Cells(1, 8), Cells(UnmatchedRow - 1, 10)).Sort Key1:=.Range(Cells(2, 9), Cells(UnmatchedRow - 1, 9)), Order1:=xlAscending, Header:=xlYes
    
    'Calculate amount sum
    .Range("L2").Formula = "=Sum(C:C)"
    IgnoreGreenTriangle .Range("L2")
    
    'Conditional highlighting for amount sum and total gross sum (if equal, green, if not equal, red)
    If Abs(.Range("L2").Value - .Range("M2").Value) < 0.01 Then
        .Range("L2").Interior.Color = 14610923
        .Range("M2").Interior.Color = 14610923
    Else
        .Range("L2").Interior.Color = 255
        .Range("M2").Interior.Color = 255
    End If
    
    'Format nicely and autofit columns
    .Columns("E").NumberFormat = "yyyy-mm-dd;@"
    .Columns("I").NumberFormat = "yyyy-mm-dd;@"
    .Columns("F").NumberFormat = "0.00;@"
    .Columns("J").NumberFormat = "0.00;@"
    .Columns("L").NumberFormat = "0.00;@"
    .Columns("D").Select
    .Columns("D").AutoFit
    .Columns("E").Select
    .Columns("E").AutoFit
    .Columns("H").Select
    .Columns("H").AutoFit
    .Columns("I").Select
    .Columns("I").AutoFit
    .Columns("L").Select
    .Columns("L").AutoFit
    
    'Final highlighting (if using just stripe column highlighting)
    If HighlightingState Then
        'do nothing, highlighting is complete
    Else
        'Highlight column C according to rules based on column D (whether a match was found, if it's a single, double, or triple)
        For j = 2 To FinalRow
            RowString = CStr(j)
            'Check if column D is empty
            If IsEmpty(.Range("D" + RowString)) Then
                If InStr(1, .Range("A" + RowString).Value, "Monthly Stripe Adjustment - ") <> 1 Then
                    'No Match Found - highlight with red
                    .Range("C" + RowString).Interior.Color = 255
                Else
                    'No Match Found - but it's for a Monthly Stripe Adjustment - highlight with light blue
                    .Range("C" + RowString).Interior.Color = 15849925
                End If
            Else
                'A Match is Found! Check if it is a "
                If .Range("D" + RowString).Value = """" Then
                    'It's a double match! Check if it's a triple
                    RowString2 = CStr(j - 1)
                    If .Range("D" + RowString2).Value = """" Then
                        'It's a triple! - highlight with dark green
                        RowString3 = CStr(j - 2)
                        .Range("C" + RowString).Interior.Color = 3969910
                        .Range("C" + RowString2).Interior.Color = 3969910
                        .Range("C" + RowString3).Interior.Color = 3969910
                    Else
                        'It's a double! - highlight with medium green
                        .Range("C" + RowString).Interior.Color = 10213316
                        .Range("C" + RowString2).Interior.Color = 10213316
                    End If
                Else
                    'It's a single match - highlight with light green
                    .Range("C" + RowString).Interior.Color = 14610923
                End If
            End If
        Next j
    End If
    
    'Final Selection and bold headers
    .Range("A1").Select
    .Rows(1).Font.Bold = True
End With

End Sub

Private Sub SearchPreviouslyUnmatchedTransactions(CurrentWorksheet As Worksheet, CurrentMonthRow As Long, DateTolerance As Long)
    Dim FoundRow As Long
    Dim RowString As Long
    Dim anotherstring As String
    
    On Error Resume Next
    FoundRow = 0
    FoundRow = Application.WorksheetFunction.Match(CDbl(Format(CurrentWorksheet.Range("C" + CStr(CurrentMonthRow)).Value, "0.00")), CurrentWorksheet.Range("K2:K" + CStr(CurrentWorksheet.Cells(CurrentWorksheet.Rows.Count, "K").End(xlUp).Row)), 0)
    On Error GoTo 0

    If FoundRow > 0 Then
        'Matched a transaction amount -- still need to do error handling here
        'Check that date is within the date tolerance
        anotherstring = CStr(FoundRow + 1)
        If Abs(DateDiff("d", CurrentWorksheet.Range("J" + anotherstring).Value, CurrentWorksheet.Range("B" + CStr(CurrentMonthRow)).Value)) <= DateTolerance Then
            'Found a match! Cut and paste match into matched columns
            CurrentWorksheet.Range("E" + CStr(CurrentMonthRow)).Value = CurrentWorksheet.Range("I" + anotherstring).Value
            CurrentWorksheet.Range("F" + CStr(CurrentMonthRow)).Value = CurrentWorksheet.Range("J" + anotherstring).Value
            CurrentWorksheet.Range("G" + CStr(CurrentMonthRow)).Value = CurrentWorksheet.Range("K" + anotherstring).Value
            CurrentWorksheet.Range("I" + anotherstring).Value = ""
            CurrentWorksheet.Range("J" + anotherstring).Value = ""
            CurrentWorksheet.Range("K" + anotherstring).Value = ""
        End If
    End If

End Sub

Private Sub TTC_DMY()
    Selection.TextToColumns Destination:=Selection, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 4)
End Sub

Private Sub IgnoreGreenTriangle(TargetRange As Range)
'Remove that annoying green triangle in upper right hand corner of the cells in a given range
    Dim rngCell As Range, bError As Byte
    For Each rngCell In TargetRange.Cells
        For bError = 1 To 7 Step 1
            With rngCell
                If .Errors(bError).Value Then
                    .Errors(bError).Ignore = True
                End If
            End With
        Next bError
    Next rngCell
End Sub

Private Sub Stripe_Merchant_Account()
'Stripe Merchant Account
'Run this macro when you have a folder containing stripe .csvs (one per month)

Dim UserInput As Integer
UserInput = MsgBox("Do you want to select a folder of stripe .csvs to process?", vbYesNoCancel, "Stripe Macro", , "Yes -- choose folder. No -- run on a single .csv that is open")

If UserInput = vbNo Then
    'Run on single csv
ElseIf UserInput = vbYes Then
    'chose a folder
ElseIf UserInput = vbCancel Then
    Exit Sub
End If

'Right now only works when have a single .csv open -- need to make it work on a folder of csvs
Dim GrossColumn As Long
Dim AmountColumn As Long
Dim FeeColumn As Long
Dim CreatedColumn As Long
Dim FinalRow As Long
Dim FinalColumn As String
Dim GrossRange As Range
Dim AmountRange As Range
Dim CreatedRange As Range
Dim FeeRange As Range
Dim StripeWorkbook As Workbook
Dim CurrentWorkbook As Workbook
Dim CurrentWorkbookPath As String
Dim CurrentWorkbookName As String
Dim CurrentRow As Long
Dim MonthStartRow As Long
Dim MonthEndRow As Long
Dim CurrentMonth As Long

'Variables for doing many .csvs at once
Dim FileTotal As Long
Dim FileIndex As Long
Dim MySplit As Variant

'Turn off screen updating and display alerts
Application.ScreenUpdating = False
Application.DisplayAlerts = False
DoEvents

'Initialize final row/column variables
FinalRow = LastRow(ActiveSheet)
FinalColumn = ConvertToLetter(LastCol(ActiveSheet))

'Define columns
GrossColumn = FindColNumber(ActiveSheet, "Total Gross", "A", FinalColumn)
AmountColumn = FindColNumber(ActiveSheet, "Total", "A", FinalColumn)
FeeColumn = FindColNumber(ActiveSheet, "Total Fees", "A", FinalColumn)
CreatedColumn = FindColNumber(ActiveSheet, "Date", "A", FinalColumn)
If CreatedColumn <> 0 Then
    'Do nothing, date column is found!
Else
    'No created date, search for "Created (UTC)"
    CreatedColumn = FindColNumber(ActiveSheet, "Created (UTC)", "A", FinalColumn)
End If

'Define ranges
Set GrossRange = ActiveSheet.Range(Cells(2, GrossColumn), Cells(FinalRow, GrossColumn))
Set AmountRange = ActiveSheet.Range(Cells(2, AmountColumn), Cells(FinalRow, AmountColumn))
Set CreatedRange = ActiveSheet.Range(Cells(2, CreatedColumn), Cells(FinalRow, CreatedColumn))
Set FeeRange = ActiveSheet.Range(Cells(2, FeeColumn), Cells(FinalRow, FeeColumn))

'Get current workbook path and name
Set CurrentWorkbook = ActiveWorkbook
CurrentWorkbookPath = ActiveWorkbook.Path
CurrentWorkbookName = ActiveWorkbook.Name

'Create new workbook
Set StripeWorkbook = Workbooks.Add
StripeWorkbook.Activate

'Column headers
Range("A1").Value = "Date"
Range("B1").Value = "Total Gross"
Range("C1").Value = "Total Fees"
Range("D1").Value = "Total"
Range("E1").Value = "Sum"
Range("F1").Value = "Fee"
Range("A1:F1").Font.Bold = True

'Copy columns into new workbook
CreatedRange.Copy
Range("A2").PasteSpecial xlPasteValues
GrossRange.Copy
Range("B2").PasteSpecial xlPasteValues
FeeRange.Copy
Range("C2").PasteSpecial xlPasteValues
AmountRange.Copy
Range("D2").PasteSpecial xlPasteValues

'Autofit and formatcolumns
Columns("A").Select
Columns("A").AutoFit
Columns("A").NumberFormat = "yyyy-mm-dd;@"
Columns("B").Select
Columns("B").AutoFit
Columns("B").NumberFormat = "0.00;@"
Columns("C").Select
Columns("C").AutoFit
Columns("C").NumberFormat = "0.00;@"
Columns("D").Select
'Columns("D").AutoFit 'don't autofit amount, gets too small!
Columns("D").NumberFormat = "0.00;@"

'Calculate Sum and Fee Sums
CurrentRow = 2
MonthStartRow = 2
CurrentMonth = Month(Cells(CurrentRow, CreatedColumn).Value)
Do While CurrentRow <= FinalRow
    'Find End of Month Row
    If Month(Cells(CurrentRow, CreatedColumn).Value) <> CurrentMonth Then
        MonthEndRow = CurrentRow - 1
        Range("E" + CStr(MonthStartRow)).Formula = "=sum(D" + CStr(MonthStartRow) + ":D" + CStr(MonthEndRow) + ")"
        Range("F" + CStr(MonthStartRow)).Formula = "=sum(C" + CStr(MonthStartRow) + ":C" + CStr(MonthEndRow) + ")"
        MonthStartRow = CurrentRow
        CurrentMonth = Month(Cells(CurrentRow, CreatedColumn).Value)
    Else
        'Same month, continue
    End If
    CurrentRow = CurrentRow + 1
Loop

'Final Month
MonthEndRow = FinalRow
Range("E" + CStr(MonthStartRow)).Formula = "=sum(D" + CStr(MonthStartRow) + ":D" + CStr(MonthEndRow) + ")"
Range("F" + CStr(MonthStartRow)).Formula = "=sum(C" + CStr(MonthStartRow) + ":C" + CStr(MonthEndRow) + ")"
        

'Last Selection
Range("A1").Select

'Turn off screen updating and display alerts
Application.ScreenUpdating = True
Application.DisplayAlerts = True
DoEvents

'Save workbook as an xlsx instead of csv
On Error Resume Next
ActiveWorkbook.SaveAs CurrentWorkbookPath + ":" + Replace(CurrentWorkbookName, ".csv", ".xlsx"), FileFormat:=52

End Sub
