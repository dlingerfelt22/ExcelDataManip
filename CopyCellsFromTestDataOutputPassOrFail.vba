Rem Attribute VBA_ModuleType=VBADocumentModule
Option VBASupport 1
Private Sub CommandButton1_Click()
    'caffeine equals code but it can be bad code any issues ask dlingerfelt22
    Dim SummarySheet As Worksheet
    Dim FolderPath As String
    Dim NRow As Long
    Dim FileName As String
    Dim WorkBk As Workbook
    Dim SourceRange As Range
    Dim DestRange As Range
    
    'speeds up process by disabling visuals and calculations
    With Application
        CalcMode = .Calculation
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
    End With
    
    ' Create a new workbook and set a variable to the first sheet.
    Set SummarySheet = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    
    ' Modify this folder path to point to the files you want to use. ADD a \ at the end
    FolderPath = "\\A-folder-path\" 'DONT FORGET YOUR FOLDER PATHS END IN \
    ' NRow keeps track of where to insert new rows in the destination workbook.
    NRow = 1
    
    ' Call Dir the first time, pointing it to all Excel files in the folder path.
    FileName = Dir(FolderPath & "*.xls*")
    
    ' Loop until Dir returns an empty string.
    Do While FileName <> ""
    '
    '
    '
        ' Open a workbook in the folder
        Set WorkBk = Workbooks.Open(FolderPath & FileName)
        
        ' Set the cell in column A to be the file name.
        SummarySheet.Range("A" & NRow).Value = FileName
        
        ' Set the source range to be range required
        ' Modify this range for your workbooks.
        ' It can span multiple rows.
            'CHANGE THIS TO FIT FILE TEST NAMES
            'checked to see if its the
            'Below we look for specific tests, edit them to make it work for your test. In this example each test has a location on the combined sheet. Leave the wild cards, I bet the name is not perfect.
    If FileName Like "*Foxtrot*" Then
        Set SourceRange = WorkBk.Worksheets(1).Range("B19:E20")
        'check to see if its a
        'Foxtrot flow test
    ElseIf FileName Like "*Tacos*" Then
        Set SourceRange = WorkBk.Worksheets(1).Range("B18:D19")
        'Check to see if its a
        'Tacos test
    ElseIf FileName Like "*Bananas*" Then
        Set SourceRange = WorkBk.Worksheets(1).Range("B18:D19")
        'check to see if its a
        'Bananas leak test
    ElseIf FileName Like "*I-Bet-You-Sang-That-Like-The-Popstar-Its-okay-we-all-did*" Then
        Set SourceRange = WorkBk.Worksheets(1).Range("B18:D19")
        'check to see if its a
        'I-Bet-You-Sang-That-Like-The-Popstar-Its-okay-we-all-did test
    ElseIf FileName Like "*Porsche*" Then
        Set SourceRange = WorkBk.Worksheets(1).Range("B23:E25")
        'check to see if its a
        'Porsche test
    Else: Set SourceRange = WorkBk.Worksheets(1).Range("A1:A1")
    End If
    
    
        ' Set the destination range to start at column B and
        ' be the same size as the source range.
        
        Set DestRange = SummarySheet.Range("B" & NRow)
        Set DestRange = DestRange.Resize(SourceRange.Rows.Count, _
           SourceRange.Columns.Count)
           
        ' Copy over the values from the source to the destination.
        DestRange.Value = SourceRange.Value
        ' Increase NRow so that we know where to copy data next.
        NRow = NRow + DestRange.Rows.Count + 1
        ' Close the source workbook without saving changes.
        WorkBk.Close savechanges:=False
        
        ' Use Dir to get the next file name.
        FileName = Dir()
        

    Loop
    
    ' Call AutoFit on the destination sheet so that all
    ' data is readable now.
    'below we test cells to see if they are equal to a failure or not. Correct this based on your testing purposes.
    SummarySheet.Columns.AutoFit
    SummarySheet.Range("A1:F8900").FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""fail"""
    SummarySheet.Range("A1:F8900").FormatConditions(1).Interior.ColorIndex = 3
    SummarySheet.Range("A1:F8900").FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""pass"""
    SummarySheet.Range("A1:F8900").FormatConditions(2).Interior.ColorIndex = 4
    
    'returns application to normal
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = CalcMode
    End With
End Sub
