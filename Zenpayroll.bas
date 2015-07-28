Attribute VB_Name = "Zenpayroll"
Private Sub Zenpayroll_Journal_Report()
'Make sure that it is a zenpayroll journal report that is open

If InStr(1, ActiveSheet.Range("A1").Value, "Payroll Journal Report") > 0 Then
    'It is a Zenpayroll Journal Report
    Zenpayroll_Journal_Report_Macro
Else
    'It is not a Zenpayroll Journal Report
    MsgBox "Zenpayroll macro runs on Payroll Journal Reports"
End If

End Sub

Private Sub Zenpayroll_Journal_Report_Macro()
'Run this macro when you have a zenpayroll journal report csv open
'This macro automatically hides unnecessary columns and then calculates
'EmployeeTax, Employer Tax, payroll tax expense and payroll tax payable for each pay day in the report

'Variables
Dim paydayRow As Long
Dim paydayCounter As Long
Dim CurrentRow As Long
Dim FinalRow As Long
Dim FinalColumn As Long
Dim payrollRow As Long
Dim GrossTotalAddress As String
Dim NetPayTotalAddress As String
Dim EmployerCostTotalAddress As String
Dim MedicalEmployeeAddress1 As String
Dim MedicalEmployeeAddress2 As String
Dim MedicalEmployerAddress1 As String
Dim MedicalEmployerAddress2 As String
Dim ColorIndexConstant As Long
Dim HasMedical As Boolean
Dim FinalFileName As String

'Initialize row variables
CurrentRow = 1
FinalRow = LastRow(ActiveSheet)
FinalColumn = LastCol(ActiveSheet)
paydayCounter = 0
ColorIndexConstant = 41
HasMedical = DetermineIfHasMedical(ActiveSheet)

'Loop through total number of pay days in report
Do While CurrentRow < FinalRow
    paydayRow = FindRowNumber(ActiveSheet, "Pay day", CurrentRow, FinalRow)
    If paydayRow <> 0 Then
        'Headers for Gross and NetPay Columns
        Range("A" + CStr(paydayRow)).Offset(1, 15).Value = "Salary & Wage Expense"
        Range("A" + CStr(paydayRow)).Offset(1, 16).Value = "Withdrawn from Bank"
        Range("A" + CStr(paydayRow)).Offset(1, 15).Font.Bold = True
        Range("A" + CStr(paydayRow)).Offset(1, 16).Font.Bold = True
        Range("A" + CStr(paydayRow)).Offset(2, 15).Font.Bold = True 'Gross
        Range("A" + CStr(paydayRow)).Offset(2, 16).Font.Bold = True 'NetPay
        'Range("A" + CStr(paydayRow)).Offset(2, 19).Font.Bold = True 'EmployerCost
        
        'Headers for Payroll Tax Expense and Payable
        Range("A" + CStr(paydayRow)).Offset(1, FinalColumn + 1).Value = "Employee Tax"
        Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 1).Value = "Dr. Salary & Wage"
        Range("A" + CStr(paydayRow)).Offset(1, FinalColumn + 2).Value = "Employer Tax"
        Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 2).Value = "Dr. Payroll Tax Expense"
        Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 3).Value = "Taxes Remitted"
        If HasMedical Then
            Range("A" + CStr(paydayRow)).Offset(1, FinalColumn + 4).Value = "Medical/Dental (Employee)"
            Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 4).Value = "Dr. Salary & Wage"
            Range("A" + CStr(paydayRow)).Offset(1, FinalColumn + 5).Value = "Medical/Dental (Employer)"
            Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 5).Value = "Dr. Employee Benefits Expense"
            Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 6).Value = "Total Benefits"
        End If
        
        Range("A" + CStr(paydayRow)).Offset(1, FinalColumn + 1).Font.Bold = True
        Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 1).Font.Bold = True
        Range("A" + CStr(paydayRow)).Offset(1, FinalColumn + 2).Font.Bold = True
        Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 2).Font.Bold = True
        Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 3).Font.Bold = True
        If HasMedical Then
            Range("A" + CStr(paydayRow)).Offset(1, FinalColumn + 4).Font.Bold = True
            Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 4).Font.Bold = True
            Range("A" + CStr(paydayRow)).Offset(1, FinalColumn + 5).Font.Bold = True
            Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 5).Font.Bold = True
            Range("A" + CStr(paydayRow)).Offset(2, FinalColumn + 6).Font.Bold = True
        End If
        
        'Get payrollRow
        payrollRow = FindRowNumber(ActiveSheet, "PAYROLL", paydayRow, FinalRow)
        If payrollRow <> 0 Then
            'Update paydayCounter (reuse same five colors)
            If paydayCounter < 5 Then
                paydayCounter = paydayCounter + 1
            Else
                paydayCounter = 1
            End If
            
            'Get EmployerCost, Gross, and NetPay totals addresses
            GrossTotalAddress = Range("P" + CStr(payrollRow)).Address
            NetPayTotalAddress = Range("Q" + CStr(payrollRow)).Address
            EmployerCostTotalAddress = Range("T" + CStr(payrollRow)).Address
            
            'Dr. Salary & Wage and Payroll Tax Expense need to subtract medical if it's included
            If HasMedical Then
                'Get Medical/Dental Employee/Employer addresses
                MedicalEmployeeAddress1 = Range("AB" + CStr(payrollRow)).Address
                MedicalEmployeeAddress2 = Range("AD" + CStr(payrollRow)).Address
                MedicalEmployerAddress1 = Range("AC" + CStr(payrollRow)).Address
                MedicalEmployerAddress2 = Range("AE" + CStr(payrollRow)).Address
                
                'Dr. Salary & Wage (Employee Tax) = Taxes Remitted - Dr. Payroll Tax Expense = Gross - Net Pay - Medical Employee
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 1).Formula = "=" + GrossTotalAddress + "-" + NetPayTotalAddress + "-(" + MedicalEmployeeAddress1 + "+" + MedicalEmployeeAddress2 + ")"
                           
                'Dr. Payroll Tax Expense = Employer Cost - Gross - Medical Employer
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 2).Formula = "=" + EmployerCostTotalAddress + "-" + GrossTotalAddress + "-(" + MedicalEmployerAddress1 + "+" + MedicalEmployerAddress2 + ")"
    
                'Medical Employee
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 4).Formula = "=" + MedicalEmployeeAddress1 + "+" + MedicalEmployeeAddress2
                 
                'Medical Employer
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 5).Formula = "=" + MedicalEmployerAddress1 + "+" + MedicalEmployerAddress2
                 
                 'Total Benefits
                 Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 6).Formula = "=SUM(" + MedicalEmployeeAddress1 + ":" + MedicalEmployerAddress2 + ")"
            
                'Taxes Remitted = Employer Cost - Net Pay - Medical Employee - Medical Employer
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 3).Formula = "=" + EmployerCostTotalAddress + "-" + NetPayTotalAddress + "-" + MedicalEmployeeAddress1 + "-" + MedicalEmployeeAddress2 + "-" + MedicalEmployerAddress1 + "-" + MedicalEmployerAddress2
            Else
                'Dr. Salary & Wage (Employee Tax) = Taxes Remitted - Dr. Payroll Tax Expense = Gross - Net Pay
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 1).Formula = "=" + GrossTotalAddress + "-" + NetPayTotalAddress
                           
                'Dr. Payroll Tax Expense = Employer Cost - Gross
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 2).Formula = "=" + EmployerCostTotalAddress + "-" + GrossTotalAddress
                
                'Taxes Remitted = Employer Cost - Net Pay
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 3).Formula = "=" + EmployerCostTotalAddress + "-" + NetPayTotalAddress
            End If
            
            'Highlights (comment out if not wanted!)
            Range("A" + CStr(paydayRow)).Offset(0, 1).Interior.ColorIndex = ColorIndexConstant + paydayCounter 'PayDay cell
            Range(GrossTotalAddress).Interior.ColorIndex = ColorIndexConstant + paydayCounter 'Gross Total
            Range(NetPayTotalAddress).Interior.ColorIndex = ColorIndexConstant + paydayCounter 'NetPay Total
            Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 1).Interior.ColorIndex = ColorIndexConstant + paydayCounter
            Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 2).Interior.ColorIndex = ColorIndexConstant + paydayCounter
            Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 3).Interior.ColorIndex = ColorIndexConstant + paydayCounter
            If HasMedical Then
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 4).Interior.ColorIndex = ColorIndexConstant + paydayCounter
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 5).Interior.ColorIndex = ColorIndexConstant + paydayCounter
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 6).Interior.ColorIndex = ColorIndexConstant + paydayCounter
            End If
            
            'Date text becomes excel date
            Range("A" + CStr(paydayRow)).Offset(0, 1).Select
            Selection.TextToColumns _
                Destination:=Selection, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
                ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, _
                Space:=False, Other:=True, FieldInfo:=Array(1, 3)
            
            'NumberFormat for highlighted values
            Range("A" + CStr(paydayRow)).Offset(0, 1).NumberFormat = "yyyy-mm-dd;@"
            Range(GrossTotalAddress).NumberFormat = "0.00"
            Range(NetPayTotalAddress).NumberFormat = "0.00"
            Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 1).NumberFormat = "0.00"
            Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 2).NumberFormat = "0.00"
            Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 3).NumberFormat = "0.00"
            If HasMedical Then
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 4).NumberFormat = "0.00"
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 5).NumberFormat = "0.00"
                Range("A" + CStr(paydayRow)).Offset(3, FinalColumn + 6).NumberFormat = "0.00"
            End If
            
            'Bench color Highlights
            'Range("A" + CStr(paydayRow)).Offset(0, 1).Interior.Color = 2723283 'Bench gold=2723283
            'Range("A" + CStr(paydayRow)).Offset(3, finalColumn + 1).Interior.Color = 5073908 'Orange
            'Range("A" + CStr(paydayRow)).Offset(3, finalColumn + 2).Interior.Color = 13101755 'Light green
            'Range(Range("A" + CStr(paydayRow)).Offset(3, 15).Address + ":" + GrossTotalAddress).Interior.Color = 5073908 'Gross=Bench orange
            'Range(Range("A" + CStr(paydayRow)).Offset(3, 16).Address + ":" + NetPayTotalAddress).Interior.Color = 13101755 'NetPay=Bench light green
            'Range(Range("A" + CStr(paydayRow)).Offset(3, 19).Address + ":" + EmployerCostTotalAddress).Interior.Color = 7700278 'EmployerCost=Bench green
            
            
            'Update currentRow
            CurrentRow = payrollRow + 1
        Else
            Exit Do
        End If
    Else
        Exit Do
    End If
Loop

'Hide columns
ZenpayrollHideColumns FinalColumn

'Autofit columns
Columns("A").Select
Columns("A").AutoFit
Columns("B").Select
Columns("B").AutoFit
Columns("C").Select
Columns("C").AutoFit
Columns("P").Select
Columns("P").AutoFit
Columns("Q").Select
Columns("Q").AutoFit

Columns(ColumnName_gibbo(Cells(1, FinalColumn + 2))).Select
Columns(ColumnName_gibbo(Cells(1, FinalColumn + 2))).AutoFit
Columns(ColumnName_gibbo(Cells(1, FinalColumn + 3))).Select
Columns(ColumnName_gibbo(Cells(1, FinalColumn + 3))).AutoFit
Columns(ColumnName_gibbo(Cells(1, FinalColumn + 4))).Select
Columns(ColumnName_gibbo(Cells(1, FinalColumn + 4))).AutoFit
If HasMedical Then
    Columns(ColumnName_gibbo(Cells(1, FinalColumn + 5))).Select
    Columns(ColumnName_gibbo(Cells(1, FinalColumn + 5))).AutoFit
    Columns(ColumnName_gibbo(Cells(1, FinalColumn + 6))).Select
    Columns(ColumnName_gibbo(Cells(1, FinalColumn + 6))).AutoFit
    Columns(ColumnName_gibbo(Cells(1, FinalColumn + 7))).Select
    Columns(ColumnName_gibbo(Cells(1, FinalColumn + 7))).AutoFit
End If

'Last Selection
Range("A1").Select

'Save workbook as an xlsx instead of csv
FinalFileName = ActiveWorkbook.Path + ":" + Replace(ActiveWorkbook.Name, ".csv", ".xlsx")
On Error Resume Next
ActiveWorkbook.SaveAs FinalFileName, FileFormat:=52

End Sub


Private Sub ZenpayrollHideColumns(FinalColumn As Long)
Attribute ZenpayrollHideColumns.VB_ProcData.VB_Invoke_Func = " \n14"
' Hide appropriate columns in Zenpayroll Payroll Journal Report .csv
    Columns("D:O").Select
    Selection.EntireColumn.Hidden = True
    Columns("R:T").Select
    Selection.EntireColumn.Hidden = True
    'Columns("U:AL").Select
    'Make final selection programatic
    If FinalColumn > 20 Then
        Range(Cells(1, 21), Cells(1, FinalColumn)).Select
        Selection.EntireColumn.Hidden = True
    End If
End Sub

Private Function ColumnName_gibbo(cellRange As Range) As String
    ColumnName_gibbo = Split(cellRange.Address(1, 0), "$")(0)
End Function

Private Function DetermineIfHasMedical(WS As Worksheet) As Boolean
    If WS.Range("AB10").Value = "Dental Insurance (Pre-Tax EE)" And WS.Range("AC10").Value = "Dental Insurance (Pre-Tax ER)" And _
       WS.Range("AD10").Value = "Medical Insurance (Pre-Tax EE)" And WS.Range("AE10").Value = "Medical Insurance (Pre-Tax ER)" Then
        DetermineIfHasMedical = True
    Else
        DetermineIfHasMedical = False
    End If
End Function

Private Function FindRowNumber(CurrentWorksheet As Worksheet, findText As String, rowStart As Long, rowEnd As Long) As Long
    Dim FindRow As Range
            
    Set FindRow = CurrentWorksheet.Range("A" + CStr(rowStart) + ":A" + CStr(rowEnd)).Find(What:=findText, LookIn:=xlValues, MatchCase:=True)
    
    If FindRow Is Nothing Then
        FindRowNumber = 0
    Else
        FindRowNumber = FindRow.Row
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
        MsgBox "Error " & Err.Number & ": " & Err.Description, _
               vbExclamation, "LastRow()"
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
        MsgBox "Error " & Err.Number & ": " & Err.Description, _
               vbExclamation, "LastRow()"
        Resume ErrExit

End Function
