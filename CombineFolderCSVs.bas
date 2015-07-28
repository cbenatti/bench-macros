Attribute VB_Name = "CombineFolderCSVs"
Option Explicit

'Important: this Dim line must be at the top of your module
Dim MyFiles As String

Private Sub CombineCSVsInUserDefinedFolder()
'Macro for having the user select a folder containing csvs with headers
'The first csv's header will be kept, the following only data will be copied in

'----------------------------------------------------------------------'
'Clear MyFiles to be sure that it not return old info if no files are found
MyFiles = ""

'----------------------------------------------------------------------'
'Define Variable names
DoEvents
Dim InputDirectory As String
Dim CurrentFileName As String
Dim CurrentFileWorkbook As Workbook
Dim FileTotal As Long
Dim FileIndex As Long
Dim MySplit As Variant
Dim SummaryReportWorkbook As Workbook
Dim SummaryReportName As String
Dim ReportFinalRow As Long
Dim CurrentFinalRow As Long
Dim CurrentFinalColumn As Long

'Initialize summary report name
SummaryReportName = "CSVs Combined.csv"

'Get input file directory
InputDirectory = Select_Folder_On_Mac("Choose folder containing all CSVs to combine")
'Check that a folder was selected
If FileOrFolderExistsOnMac(0, InputDirectory) Then
    'Valid folder selected, continue -- get all CSVs within folder
    Call GetFilesOnMacWithOrWithoutSubfolders(FolderPath:=InputDirectory, Level:=1, ExtChoice:=5, FileFilterOption:=0, FileNameFilterStr:="")

    If MyFiles <> "" Then
        
        'Turn off screen updating and display alerts
        With Application
            .ScreenUpdating = False
            .DisplayAlerts = False
        End With
        
        'Get file names to open from MyFiles
        MySplit = Split(MyFiles, Chr(10))
        
        'Do stuff with the files if MyFiles is not empty
        FileTotal = UBound(MySplit)
        
        'First csv file -- keep header
        On Error Resume Next
        CurrentFileName = MySplit(FileIndex)
        
        'Open First File
        Workbooks.Open CurrentFileName
        Set CurrentFileWorkbook = ActiveWorkbook
        On Error Resume Next
        ActiveWorkbook.SaveAs ActiveWorkbook.Path + ":" + SummaryReportName, FileFormat:=6

        'Define ReportWorkbook and final row
        Set SummaryReportWorkbook = ActiveWorkbook
        ReportFinalRow = LastRow(SummaryReportWorkbook.Worksheets(1))

        On Error GoTo 0
        
        For FileIndex = LBound(MySplit) + 1 To FileTotal - 1
            On Error Resume Next
            CurrentFileName = MySplit(FileIndex)
            
            'Open Income Statement
            Workbooks.Open CurrentFileName
            Set CurrentFileWorkbook = ActiveWorkbook
            
            'Define final row and column
            CurrentFinalRow = LastRow(CurrentFileWorkbook.Worksheets(1))
            CurrentFinalColumn = LastCol(CurrentFileWorkbook.Worksheets(1))
            
            'Copy and paste into report workbook
            CurrentFileWorkbook.Worksheets(1).Range(Cells(2, 1), Cells(CurrentFinalRow, CurrentFinalColumn)).Copy _
                Destination:=SummaryReportWorkbook.Worksheets(1).Range("A" + CStr(ReportFinalRow + 1))
            ReportFinalRow = LastRow(SummaryReportWorkbook.Worksheets(1))
            
            'Close Workbook
            CurrentFileWorkbook.Close
            DoEvents

            On Error GoTo 0
        Next FileIndex
        
        'Turn on screen updating and display alerts
        With Application
            .ScreenUpdating = True
            .DisplayAlerts = True
        End With
        
        'Save workbook, and final selection
        SummaryReportWorkbook.Activate
        SummaryReportWorkbook.Worksheets(1).Range("A1").Select
        SummaryReportWorkbook.Save
        
    Else
        MsgBox "Sorry no files that match your criteria"
        
        'ScreenUpdating is still True but we set it to true again to refresh the screen,
        With Application
            .ScreenUpdating = True
        End With
    End If

Else
    'No folder selected, therefore exit macro
    Exit Sub
End If
DoEvents

'----------------------------------------------------------------------'
'Turn on screen updatingand display messages
Application.ScreenUpdating = True
Application.DisplayAlerts = True
    
'----------------------------------------------------------------------'

End Sub

Private Function Select_Folder_On_Mac(PromptMessage As String)
    Dim FolderPath As String
    Dim RootFolder As String

    On Error Resume Next
    'Start at desktop folder
    RootFolder = MacScript("return (path to desktop folder) as String")
    'For testing, start at predefined directory
    'RootFolder = "Macintosh HD:Users:benchemployee:Desktop:Excel:TaxPackage:TN_Financial_Documents:2014_Tip_Network_Financials:"
    FolderPath = MacScript("(choose folder with prompt """ + PromptMessage + """" & _
    "default location alias """ & RootFolder & """) as string")
    On Error GoTo 0

    Select_Folder_On_Mac = FolderPath
    
End Function

Private Function FileOrFolderExistsOnMac(FileOrFolder As Long, FileOrFolderstr As String) As Boolean
'By Ron de Bruin
'30-July-2012
'Function to test whether a file or folder exist on a Mac.
'Uses AppleScript to avoid the problem with long file names.
    Dim ScriptToCheckFileFolder As String
    ScriptToCheckFileFolder = "tell application " & Chr(34) & "Finder" & Chr(34) & Chr(13)
    If FileOrFolder = 1 Then
        ScriptToCheckFileFolder = ScriptToCheckFileFolder & "exists file " & _
        Chr(34) & FileOrFolderstr & Chr(34) & Chr(13)
    Else
        ScriptToCheckFileFolder = ScriptToCheckFileFolder & "exists folder " & _
        Chr(34) & FileOrFolderstr & Chr(34) & Chr(13)
    End If
    ScriptToCheckFileFolder = ScriptToCheckFileFolder & "end tell" & Chr(13)
    FileOrFolderExistsOnMac = MacScript(ScriptToCheckFileFolder)
End Function

'*******Function that does all the hard work that will be called by the macro*********

Private Function GetFilesOnMacWithOrWithoutSubfolders(FolderPath As String, Level As Long, ExtChoice As Long, _
                                              FileFilterOption As Long, FileNameFilterStr As String)
'Ron de Bruin,Version 2.2: 1 Jan 2015
'http://www.rondebruin.nl/mac.htm
'Thanks to DJ Bazzie Wazzie(poster on MacScripter) for his great help.
    Dim ScriptToRun As String
    'Dim folderPath As String
    Dim FileNameFilter As String
    Dim Extensions As String

    On Error Resume Next
    'folderPath = MacScript("choose folder as string")
    If FolderPath = "" Then Exit Function
    On Error GoTo 0

    Select Case ExtChoice
    Case 0: Extensions = "(xls|xlsx|xlsm|xlsb)"  'xls, xlsx , xlsm, xlsb
    Case 1: Extensions = "xls"    'Only  xls
    Case 2: Extensions = "xlsx"    'Only xlsx
    Case 3: Extensions = "xlsm"    'Only xlsm
    Case 4: Extensions = "xlsb"    'Only xlsb
    Case 5: Extensions = "csv"    'Only csv
    Case 6: Extensions = "txt"    'Only txt
    Case 7: Extensions = ".*"    'All files with extension, use *.* for everything
    Case 8: Extensions = "(xlsx|xlsm|xlsb)"  'xlsx, xlsm , xlsb
    Case 9: Extensions = "(csv|txt)"   'csv and txt files
    Case 10: Extensions = "CSV" 'Capital CSV
        'You can add more filter options if you want,
    End Select

    Select Case FileFilterOption
    Case 0: FileNameFilter = "'.*/[^~][^/]*\\." & Extensions & "$' "  'No Filter
    Case 1: FileNameFilter = "'.*/" & FileNameFilterStr & "[^~][^/]*\\." & Extensions & "$' "    'Begins with
    Case 2: FileNameFilter = "'.*/[^~][^/]*" & FileNameFilterStr & "\\." & Extensions & "$' "    ' Ends With
    Case 3: FileNameFilter = "'.*/([^~][^/]*" & FileNameFilterStr & "[^/]*|" & FileNameFilterStr & "[^/]*)\\." & Extensions & "$' "   'Contains
    End Select

    FolderPath = MacScript("tell text 1 thru -2 of " & Chr(34) & FolderPath & _
                           Chr(34) & " to return quoted form of it's POSIX Path")
    FolderPath = Replace(FolderPath, "'\''", "'\\''")
    ScriptToRun = ScriptToRun & _
                  "set streamEditorCommand to " & _
                  Chr(34) & " |  tr  [/:] [:/] " & Chr(34) & Chr(13)
    ScriptToRun = ScriptToRun & _
                  "set streamEditorCommand to streamEditorCommand & " & _
                  Chr(34) & " | sed -e " & Chr(34) & "  & quoted form of (" & _
                  Chr(34) & " s.:." & Chr(34) & _
                "  & (POSIX file " & Chr(34) & "/" & Chr(34) & "  as string) & " & _
                  Chr(34) & "." & Chr(34) & " )" & Chr(13)
    ScriptToRun = ScriptToRun & "do shell script """ & "find -E " & _
                  FolderPath & " -iregex " & FileNameFilter & "-maxdepth " & _
                  Level & """ & streamEditorCommand without altering line endings"

    On Error Resume Next
    MyFiles = MacScript(ScriptToRun)
    On Error GoTo 0
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
