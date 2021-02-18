Option Explicit

Private Sub DoSimpleReplacements(ByRef WordDoc)
    Dim Finder: Set Finder = WordDoc.Content
    Finder.Find.Wrap = wdFindStop
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

Sub MaakOfferte()
    Dim MainSheet: Set MainSheet = ThisWorkbook.Sheets("Main")
    Dim WordApp: Set WordApp = CreateObject("Word.Application")
    Dim WordDoc As Word.Document: Set WordDoc = WordApp.Documents.Open(Application.ActiveWorkbook.Path & "/OfferteTemplate.docx", ReadOnly:=True)

    Dim DatePart: DatePart = Format(Now, "YYYY-MM-DD")
    Dim NamePart: NamePart = MainSheet.Range("Naam").Value
    'WordDoc.SaveAs (DatePart & " - " & NamePart & " - Offerte.docx")

    DoSimpleReplacements WordDoc

    WordApp.Visible = True
    AppActivate WordApp.Caption

End Sub
