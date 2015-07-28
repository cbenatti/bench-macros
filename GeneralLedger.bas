Attribute VB_Name = "GeneralLedger"
Private Sub GeneralLedgerSort()
'There are two types of General Ledger Reports that the bookkeeper may have open
'This macro determines which report is open, and runs the appropriate sorting macro

If InStr(1, ActiveSheet.Range("A1").Value, "General ledger report for ") > 0 Then
    'It is a General Ledger Report
    If ActiveSheet.Range("A3").Value = "Date" Then
        'It is a GL downloaded with Tax Package
        GeneralLedgerTaxPackageReport
    Else
        'It is a GL downloaded directly from the app
        GeneralLedgerDownloadFromApp
    End If
Else
    'It is not a General Ledger Report
    MsgBox "GL macro runs on GL downloaded from Bench App"
End If

End Sub

Private Sub GeneralLedgerTaxPackageReport()
'This macro should be run when you have a General Ledger tab open.
'The columns are: Date, Description, Account, Dr, Cr, with data starting in row 4
'This macro creates a new sheet with a table of data able to be sorted by account, either Credit or Debit

'Read in General Ledger Data
Dim RawData() As Range
Dim FinalRow As Long
Dim GLWorkbook As Workbook
Dim GLWorksheet As Worksheet
Set GLWorkbook = ActiveWorkbook
Set GLWorksheet = Worksheets("General Ledger")
GLWorksheet.Copy after:=GLWorksheet
Set GLWorksheet = Worksheets("General Ledger (2)")

'Get max row to iterate through
FinalRow = LastRow(GLWorksheet)

'Rename header in column d to Dr/Cr
Cells(3, 4).Value = "Dr/Cr"
Cells(3, 5).Value = ""

'For each row...
Dim current_row As Long
For current_row = 4 To FinalRow Step 1
    'If column b is empty, and column c is not empty, fill column b cell with cell from column b in above row
    If IsEmpty(Cells(current_row, 2)) And Not IsEmpty(Cells(current_row, 3)) Then
        Cells(current_row, 2).Value = Cells(current_row - 1, 2)
        'Remove Carriage Returns
        Cells(current_row, 2) = WorksheetFunction.Substitute(Cells(current_row, 2), Chr(13), " ")
        Cells(current_row, 2).WrapText = False
    End If
Next current_row

'Multiply column d by -1
Range(Cells(4, 4), Cells(FinalRow, 4)).Select
MultiplybyNegativeOne

'Delete blanks in column D shifting left
Range(Cells(4, 4), Cells(FinalRow, 4)).SpecialCells(xlCellTypeBlanks).Select
Selection.Delete Shift:=xlToLeft

'Delete blank rows shifting up
Range(Cells(4, 1), Cells(FinalRow, 4)).SpecialCells(xlCellTypeBlanks).Select
Selection.Delete Shift:=xlUp

Cells(1, 1).Select

End Sub

Private Sub GeneralLedgerDownloadFromApp()
'The open worksheet should be the GL
'This macro sorts every block by transaction alphabetically (Column A)
'It leaves MIT alone, because this should be sorted by date

'The macro also converts all text dates to Excel dates (Column B)

'Turn off screen updating and display alerts
Application.ScreenUpdating = False
Application.DisplayAlerts = False
DoEvents

'Variables
Dim FinalRow As Long
Dim CurrentRow As Long
Dim OpeningBalanceRow As Long
Dim ClosingBalanceRow As Long

'Convert text dates to Excel dates
FindReplace "B:B", "June", "Jun"    'Need to replace June and July text
FindReplace "B:B", "July", "Jul"      'Because Excel for Mac ..................
TTC_DMY
DateFormat

'Initialize Row Variables
CurrentRow = 1
OpeningBalanceRow = 1
ClosingBalanceRow = 1

'Get max row to iterate through
FinalRow = LastRow(ActiveSheet)

'Loop through Ledger Blocks
Do While CurrentRow < FinalRow
    'Ledger Block starts only if Column B is not empty
    If Not IsEmpty(Cells(CurrentRow, 2)) Then
        'Define Ledger Block start and end rows
        OpeningBalanceRow = CurrentRow
        ClosingBalanceRow = FindClosingBalanceRow(OpeningBalanceRow)
        
        'Money in transit or Money in transit (outstanding) is not sorted
        If Not (Cells(CurrentRow, 1).Value = "Money in transit" Or Cells(CurrentRow, 1).Value = "Money in transit (outstanding)") Then
            'Sort data between Opening and Closing balances
            Range(Cells(OpeningBalanceRow + 1, 1), Cells(ClosingBalanceRow - 1, 4)).Sort _
            Key1:=Range(Cells(OpeningBalanceRow + 1, 1), Cells(ClosingBalanceRow - 1, 1)), Order1:=xlAscending, Header:=xlNo
        End If
        
        'CurrentRow updated to end of Ledger Block
        CurrentRow = ClosingBalanceRow
    End If
    
    'CurrentRow updated
    CurrentRow = CurrentRow + 1
Loop

'Turn on screen updating and display alerts
Application.ScreenUpdating = True
Application.DisplayAlerts = True
DoEvents

'Final Selection A1
Cells(1, 1).Select

End Sub

Private Sub FindReplace(SelectedColumn As String, FindString As String, ReplaceString As String)
    Columns(SelectedColumn).Select
    Selection.Replace What:=FindString, Replacement:=ReplaceString, LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
End Sub

Private Sub TTC_DMY()
    Selection.TextToColumns Destination:=Selection, DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 4)
End Sub

Private Sub DateFormat()
    Selection.NumberFormat = "yyyy-mm-dd;@"
End Sub

Private Function FindClosingBalanceRow(OpeningBalanceRow As Long)
    FindClosingBalanceRow = Range("E" + CStr(OpeningBalanceRow)).End(xlDown).Row
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
        MsgBox "Error " & Err.Number & ": " & Err.Description, _
               vbExclamation, "LastRow()"
        Resume ErrExit

End Function

Private Sub MultiplybyNegativeOne()
'Multiply the cells in a selected range by -1

'Declare variables
Dim ActSheet As Worksheet   'This is the active worksheet selected by the user
Dim SelRange As Range          'This is the selection of cells selected by the user
Dim TestCell As Range           'This is a cell that iterates through SelRange
Dim FinalRow As Long

'Set Active Sheet and Selection variables
Set ActSheet = ActiveSheet
Set SelRange = Selection

FinalRow = LastRow(ActSheet)

'Check if range is empty or is not numeric
For Each TestCell In SelRange.Cells
    If TestCell.Row > FinalRow Then
        Exit For
    End If
    
    If Not IsNumeric(TestCell) Or IsEmpty(TestCell) Then 'Have to have this convention for blank cells
                                                                                        'Otherwise IsNumeric evals to true for empty cells
        'Do nothing
    Else
        'Multiply by negative one
        TestCell = TestCell.Value * -1
    End If
Next TestCell

'Put back original sheet and selection
ActSheet.Select
SelRange.Select

End Sub




