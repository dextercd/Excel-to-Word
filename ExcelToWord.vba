Option Explicit

Dim Warnings As New Collection

Private Sub WriteWarning(Warning As String)
    Warnings.Add Warning
End Sub

Private Sub DisplayWarnings()
    If Warnings.Count > 0 Then
        Dim WarningText As String: WarningText = ""
        Dim W As Variant

        For Each W In Warnings
            WarningText = WarningText & ">" & W & vbCrLf
        Next

        MsgBox WarningText, vbExclamation
        Set Warnings = Nothing
    End If
End Sub

Private Sub DoSimpleReplacements(ByRef WordDoc)
    Dim Finder: Set Finder = WordDoc.Content
    Finder.Find.MatchWildcards = True
    Finder.Find.Text = "#[a-zA-Z]@#"
    Finder.Find.Execute

    While Finder.Find.Found
        Dim Keyword: Keyword = Replace(Finder.Text, "#", "")

        ' Editing a range causes .Find.Execute to break, so we use a duplicate
        Dim EditRange: Set EditRange = Finder.Duplicate
        EditRange.Text = Range(Keyword).Text

        Finder.Find.Execute
    Wend
End Sub

Private Sub CreateTable(ByRef WordDoc As Word.Document, InsertPoint As String, ExcelList As ListObject)
    Dim WordCell As Word.Cell
    Dim Row As Integer
    Dim Col As Integer
    Dim EditRange As Word.Range
    Dim WordTable As Word.Table

    Dim Finder As Word.Range: Set Finder = WordDoc.Content
    With Finder.Find
        .Wrap = wdFindStop
        .MatchWildcards = True
        .MatchCase = False
        .Text = "\<\<" & InsertPoint & "\>\>"
    End With

    Finder.Find.Execute

    While Finder.Find.Found
        ' Editing a range causes .Find.Execute to break, so we use a duplicate
        Set EditRange = Finder.Duplicate

        Set WordTable = WordDoc.Tables.Add(Range:=EditRange, _
                                           NumRows:=ExcelList.Range.Rows.Count, _
                                           NumColumns:=ExcelList.Range.Columns.Count, _
                                           AutoFitBehavior:=wdAutoFitContent)

        With WordTable
            .Style = "Grid Table 1 Light"
            .ApplyStyleColumnBands = True
            .ApplyStyleRowBands = True
            .ApplyStyleFirstColumn = False
            .ApplyStyleHeadingRows = True
            .ApplyStyleLastColumn = False
            .ApplyStyleLastRow = False
        End With

        For Row = 1 To ExcelList.Range.Rows.Count
            For Col = 1 To ExcelList.Range.Columns.Count
                Set WordCell = WordTable.Cell(Row, Col)
                WordCell.Range.Text = ExcelList.Range(Row, Col).Text
            Next
        Next

        Finder.Find.Execute
    Wend
End Sub

Sub MaakOfferte()
    Dim MainSheet As Worksheet
    Set MainSheet = ThisWorkbook.Sheets("Main")

    Dim WordApp As Word.Application
    Set WordApp = CreateObject("Word.Application")

    Dim WordDoc As Word.Document
    Set WordDoc = WordApp.Documents.Open( _
        Application.ActiveWorkbook.Path & "/OfferteTemplate.docx", _
        ReadOnly:=True)

    Dim DatePart As String: DatePart = Format(Now, "YYYY-MM-DD")
    Dim NamePart As String: NamePart = MainSheet.Range("Naam").Value
    Dim FileName As String: FileName = DatePart & " - " & NamePart & " - Offerte.docx"
    'WordDoc.SaveAs (FileName)

    DoSimpleReplacements WordDoc
    CreateTable WordDoc, "producten", MainSheet.ListObjects("Producten")

    WordApp.Visible = True
    'AppActivate WordApp.Caption

    DisplayWarnings
End Sub
