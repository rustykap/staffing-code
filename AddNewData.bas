Attribute VB_Name = "AddNewData"
Sub AddNewData()
    Dim Master As Workbook
    Dim wb As Workbook
    Dim w_s As Worksheet
    Dim datasheet As Worksheet
    Dim myPath As String
    Dim myFile As String
    Dim myExtension As String
    Dim FldrPicker As FileDialog
    Dim numrows As Integer
    Dim numcols As Integer
    Dim pasterow As Integer
    Dim startcell As Range
    
    
    'Optimize Macro Speed
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
            
    Set Master = ThisWorkbook
    Master.Activate
    DeleteDefinedNames
    
    Sheets.Add After:=Worksheets(Worksheets.Count)
    Set datasheet = ActiveSheet
    
    pasterow = 0

'Retrieve Target Folder Path From User
    Set FldrPicker = Application.FileDialog(msoFileDialogFolderPicker)

        With FldrPicker
            .Title = "Select A Folder Containg Individual Staffing Logs"
            .AllowMultiSelect = False
                If .Show <> -1 Then GoTo NextCode
                myPath = .SelectedItems(1) & "\"
        End With

'In Case of Cancel
NextCode:
    myPath = myPath
    If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
    myExtension = "*.xls*"

'Target Path with Ending Extention
    myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
    Do While myFile <> ""
    'Set variable equal to opened workbook
        Set wb = Workbooks.Open(Filename:=myPath & myFile)
    
    'Ensure Workbook has opened before moving on to next line of code
        DoEvents
    
    'Copy the visible tab containing the data and paste into the compiler file
        For Each w_s In Worksheets
            If w_s.Visible = True And Not (IsEmpty(w_s.Range("B52"))) Then
                
                w_s.Activate
                If IsEmpty(Range("B53")) Then
                    numrows = 1
                Else: numrows = Range("B52", Range("B52").End(xlDown)).Rows.Count
                End If
            
                numcols = 21
                   
                w_s.Range("B52", Range("B52").Offset(numrows - 1, numcols - 1)).Select
                Selection.Copy
                
                Master.Activate
                datasheet.Select
                
                datasheet.Range("C3").Offset(pasterow).PasteSpecial xlPasteValues
                
                w_s.Activate
                w_s.Range("D3").Select
                Selection.Copy
                Master.Activate
                datasheet.Select
                
                Set startcell = Range("A3").Offset(pasterow)
                datasheet.Range(startcell, startcell.Offset(numrows - 1)).PasteSpecial xlPasteValues
                
                w_s.Activate
                w_s.Range("D4").Select
                Selection.Copy
                Master.Activate
                datasheet.Select
                
                Set startcell = Range("B3").Offset(pasterow)
                datasheet.Range(startcell, startcell.Offset(numrows - 1)).PasteSpecial xlPasteValues
                datasheet.Range(startcell, startcell.Offset(numrows - 1)).Select
                Selection.NumberFormat = "mm/dd/yy"
                
                pasterow = pasterow + numrows
                
            End If
        Next
    
    'Do not save and Close Workbook
        wb.Close SaveChanges:=False
      
    'Ensure Workbook has closed before moving on to next line of code
        DoEvents

    'Get next file name
        myFile = Dir
    Loop
            
    datasheet.Range("D:D,F:F,H:H,J:J,L:L,N:N,P:P,R:R,T:T,V:V").Select
    Selection.Delete
            
    'copy and paste the collated data into the bottom of the database
    
    datasheet.Activate
    numrows = Range("A3", Range("A3").End(xlDown)).Rows.Count
    numcols = 13

    datasheet.Range("A3", Range("A3").Offset(numrows - 1, numcols - 1)).Select
    Selection.Copy

    Master.Sheets("Data").Activate
    
    numpasterows = Master.Sheets("Data").Range("A3", Range("A3").End(xlDown)).Rows.Count
    Sheets("Data").Range("A3").Offset(numpasterows).Select
    Selection.PasteSpecial xlPasteValues

    'Clean up the database
    Dim rng As Range
    Dim cell As Range
    
    numrowsclean = Master.Sheets("Data").Range("A3", Range("A3").End(xlDown)).Rows.Count
    Master.Sheets("Data").Range("C3", Range("C3").Offset(numrowsclean - 1, 1)).Select
    CleanData
    
    Master.Sheets("Data").Range("B3", Range("B3").Offset(numrowsclean - 1)).Select
    Selection.NumberFormat = "mm/dd/yy"
    
    'Name the data for the pivot table
    numrowsdata = Master.Sheets("Data").Range("A2", Range("A2").End(xlDown)).Rows.Count
    numcolsdata = 13
    Set dataselect = Master.Sheets("Data").Range("A2", Range("A2").Offset(numrowsdata - 1, numcolsdata - 1))
    Master.Names.Add Name:="PivotData", RefersTo:=dataselect
    
    'delete the temporary sheet holding the compiled data
    Application.DisplayAlerts = False
    datasheet.Delete
    Application.DisplayAlerts = True

    Master.Sheets("Data").Range("A1").Select
    
    ActiveWorkbook.RefreshAll

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub
    
Sub DeleteDefinedNames()
  Dim nm As Name
  For Each nm In ActiveWorkbook.Names
    nm.Delete
  Next nm
End Sub

Sub CleanData()
  Dim cell As Range
  For Each cell In Selection
      cell.Value = Replace(cell.Value, Chr(160), " ")
      cell.Value = WorksheetFunction.Trim(cell.Value)
      cell.Value = WorksheetFunction.Proper(cell.Value)
  Next cell
End Sub

