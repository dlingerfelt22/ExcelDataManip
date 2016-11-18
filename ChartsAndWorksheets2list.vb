Rem Attribute VBA_ModuleType=VBADocumentModule 'OpenOffice syntax - DELETE LINE IF USING Micro$oft Office
Option VBASupport 1 'OpenOffice syntax - DELETE LINE IF USING Micro$oft Office
Private Sub CommandButton1_Click() 'create Sub then define variables
    Dim SName As Variant
    Dim FName As Variant
    Dim FShts As Range
    Dim MyRange As Range
    Dim MyOther As Range
    Dim Ch As Chart
    Dim i As Integer
    Dim k As Integer
    Dim WB As Workbook, MyWB As Workbook
    Dim Home As Object
    Dim Sh As Worksheet
    Dim SourceRange As Range
    Dim DestRange As Range
    Dim SummarySheet As Worksheet
    Dim NRow As Long
    Dim FOName As Workbook
    Dim actsht As Worksheet
    Dim DataSheet As Worksheet
 
    Columns(1).Insert
    Columns(1).Insert
    Set SName = ThisWorkbook
        Set SummarySheet = SName.Worksheets(1)
        NRow = 1
        Set DestRange = SummarySheet.Range("A" & NRow)
        FName = Application.GetOpenFilename 'gets user to select file to open
    If FName = False Then
        Application.ScreenUpdating = False
        SName.Activate
        Columns("A:B").Select
        Selection.Delete
        Application.ScreenUpdating = True
        Exit Sub
    Else
        Application.ScreenUpdating = False
        Workbooks.Open FName
    End If
        Sheets.Add After:=Sheets(Sheets.Count) 'adds a sheet for copying names off
        Set FOName = ActiveWorkbook
    For i = 1 To Worksheets.Count
        Worksheets(Worksheets.Count).Cells(i, 1) = Worksheets(i).Name
    Next i 'loops thru to copy names of sheets to new sheet
    k = 1
    For Each Ch In Charts
        Worksheets(Worksheets.Count).Cells(k, 2) = Ch.Name
        k = k + 1
    Next Ch
    For Each Sh In Worksheets
        Sh.Select
    Next Sh
    Set actsht = ActiveSheet
        Set SourceRange = actsht.Range("A1:B65536")
        
        ' Set the destination range to start at column B and
        ' be the same size as the source range.

        Set DestRange = DestRange.Resize(SourceRange.Rows.Count, _
           SourceRange.Columns.Count)
           DestRange.Value = SourceRange.Value
           
        SName.Activate
    Set MyWB = ThisWorkbook
        Application.ScreenUpdating = True
        Application.DisplayAlerts = False
    For Each WB In Workbooks
        If WB.Name <> MyWB.Name Then WB.Close False
    Next
        Application.DisplayAlerts = True
        Application.ScreenUpdating = False
        Application.Goto Cells(Rows.Count, "A").End(xlUp)
        ActiveCell.ClearContents
        Application.ScreenUpdating = True
End Sub

Private Sub CommandButton2_Click()
    Dim MyCell As Range, MyRange As Range
    Dim YesOrNoAnswerToMessageBox As String
    Dim QuestionToMessageBox As String
    Dim file_name As Variant

    Set MyRange = Sheets(1).Range("A1") 'starts range at A1 Looks down to infinite
    Set MyRange = Range(MyRange, MyRange.End(xlDown))
    For Each MyCell In MyRange
        Sheets.Add After:=Sheets(Sheets.Count) 'creates a new worksheet
        Sheets(Sheets.Count).Name = MyCell.Value ' renames the new worksheet
    Next MyCell 'runs until MyCell is equal to nothing
        'Creates a copy So macro page can be removed
        'begins saving process here
    file_name = Application.GetSaveAsFilename( _
        FileFilter:="Excel Files,*.xls,All Files,*.*", _
        Title:="Save As File Name")    ' Get the file name
    If file_name = False Then Exit Sub     ' See if the user canceled.
    If LCase$(Right$(file_name, 4)) <> ".xls" Then ' Save the file with the new name.
        file_name = file_name & ".xls"
    End If
    ActiveWorkbook.SaveAs Filename:=file_name 'ends saving process here
    QuestionToMessageBox = "Would you like to delete macro? Requires manual save"
    YesOrNoAnswerToMessageBox = MsgBox(QuestionToMessageBox, vbYesNo, "Delete?")
        If YesOrNoAnswerToMessageBox = vbYes Then
            Application.DisplayAlerts = False
            On Error Resume Next
            Sheets("1").Delete
            Application.DisplayAlerts = True
        Else
            MsgBox "Sheet 1 not deleted, Macro is still attached"
        End If 'After this point Run nothing it has been deleted
End Sub
Sub Start()
Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub


