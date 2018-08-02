Attribute VB_Name = "mAddressTool"
Option Explicit
'References Microsoft Word Object Library
Sub PullAddresses()
    Dim wbNew As Workbook
    Dim wsNew As Worksheet
    
    Set wbNew = Workbooks.Add(1)
    Set wsNew = wbNew.Sheets(1)
    
    wsNew.Cells.NumberFormat = "@" 'Text format so any address components that happen to look like dates etc. aren't changed.
    
    Dim docToRead As Document
    
    Dim wordApp As Word.Application
    Set wordApp = New Word.Application
    
    'An initial path can be put here for example:strInitialPath:=ControlSheet.Range("InitialFolder")
    GetDocumentReference wordApp, docToRead, "Please select a Microsoft Word document to read." 

    ReadWordDocumentAddressesToSheet docToRead, wsNew
    
    docToRead.Close
    wordApp.Quit
    
    With wsNew
        .Columns(1).Insert
        .Rows(1).Insert
        .Range("A1:G1") = Array("Print", "Line 1", "Line 2", "Line 3", "Line 4", "Line5", "Line6")
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").Borders(xlEdgeBottom).Weight = xlMedium
        .Range("A1:G1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Rows.AutoFit
        .Columns.AutoFit
        .Columns(1).NumberFormat = "General" 'To allow boolean TRUE/FALSE.  Note these are "True" and "False" in a mail merge.
        .Name = "Addresses"
    End With
End Sub

'For opening a Word Document and assigning a reference to it
Sub GetDocumentReference(ByRef msWordApp As Word.Application, ByRef docRefToSet As Document, strTitle As String, _
                         Optional strInitialPath As String)
    'Setting initial path if none was specified
    If Len(strInitialPath) = 0 Then
        strInitialPath = ThisWorkbook.Path & "\"
    End If
    
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    'FileDialog options aren't set to defaults before use, so they must be defined to avoid strange issues.
    With fd
        .Title = strTitle
        .AllowMultiSelect = False
        .InitialFileName = strInitialPath
        .ButtonName = "Select"
        .Filters.Clear
        .Filters.Add "Word Document", "*.docx", 1
        .Filters.Add "Word Document (Older)", "*.doc", 2
        .Filters.Add "All Files", "*.*", 3
        .FilterIndex = 1
        
        .Show
        
        If .SelectedItems.Count = 0 Then
            Exit Sub
        End If
        
        Set docRefToSet = msWordApp.Documents.Open(Filename:=.SelectedItems.Item(1), ReadOnly:=True, Visible:=False)
    End With
End Sub

Sub ReadWordDocumentAddressesToSheet(ByRef docInput As Document, ByRef wsOutput As Worksheet)
    Dim strTempArray() As String
    Dim colTemp As New Collection
    Dim lngCurrentRow As Long
    lngCurrentRow = 1
    
    Dim lngCurrentParagraph As Long
    Dim parCurrent As Paragraph
    
    For lngCurrentParagraph = 1 To docInput.Paragraphs.Count
        Set parCurrent = docInput.Paragraphs(lngCurrentParagraph)
        If (Len(Trim(parCurrent.Range.Text)) > 1) Then 'Length of 1 means it's only a new line character.
            colTemp.Add Trim(Left(parCurrent.Range.Text, Len(parCurrent.Range.Text) - 1))
        End If
        If (colTemp.Count > 0) And ((Len(Trim(parCurrent.Range.Text)) = 1) Or (lngCurrentParagraph = docInput.Paragraphs.Count)) Then
                CollectionToArray colTemp, strTempArray
                wsOutput.Range(wsOutput.Cells(lngCurrentRow, 1), wsOutput.Cells(lngCurrentRow, UBound(strTempArray) + 1)) = strTempArray
                lngCurrentRow = lngCurrentRow + 1
        End If
    Next
End Sub

Sub CollectionToArray(ByRef colInput As Collection, ByRef strArray() As String)
    Erase strArray
    If Not (colInput.Count = 0) Then
        ReDim strArray(colInput.Count)
        Dim lngTemp As Long

        For lngTemp = 0 To colInput.Count - 1
            strArray(lngTemp) = colInput(1)
            colInput.Remove (1)
        Next
    End If
End Sub
