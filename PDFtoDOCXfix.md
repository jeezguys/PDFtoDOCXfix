Sub PDFtoDOCXFix()
    Dim doc As Document
    Dim para As Paragraph
    Dim tbl As Table
    Const wdLineSpaceDouble = 3
    Const wdFindContinue = 1
    Const wdReplaceAll = 2

Set doc = ActiveDocument

' Helper: Run basic cleanup steps once
Call BasicCleanup(doc)

' Convert all tables to plain text (tab delimited)
For Each tbl In doc.Tables
    tbl.ConvertToText Separator:=wdSeparateByTabs, NestedTables:=True
Next tbl

' Remove empty paragraphs (zero length or just spaces)
For Each para In doc.Paragraphs
    If Len(Trim(para.Range.Text)) = 1 Then ' Only paragraph mark
        para.Range.Delete
    End If
Next para

' Run basic cleanup again to catch anything new
Call BasicCleanup(doc)

MsgBox "Deep cleanup finished: tables flattened, empty paragraphs removed, formatting applied.", vbInformation

End Sub

Sub BasicCleanup(ByRef doc As Document)
    Dim para As Paragraph

With doc.Content
    ' Font and spacing
    With .Font
        .Name = "Times New Roman"
        .Size = 12
    End With
    With .ParagraphFormat
        .LineSpacingRule = wdLineSpaceDouble
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
End With

' Remove manual page breaks
With doc.Content.Find
    .ClearFormatting
    .Replacement.ClearFormatting
    .Text = "^m"
    .Replacement.Text = " "
    .Wrap = wdFindContinue
    .Execute Replace:=wdReplaceAll
End With

' Remove section breaks
With doc.Content.Find
    .Text = "^b"
    .Replacement.Text = " "
    .Execute Replace:=wdReplaceAll
End With

' Remove manual line breaks
With doc.Content.Find
    .Text = "^l"
    .Replacement.Text = " "
    .Execute Replace:=wdReplaceAll
End With

' Replace triple paragraph breaks with double
With doc.Content.Find
    .Text = "^p^p^p"
    .Replacement.Text = "^p^p"
    Do While .Execute(Replace:=wdReplaceAll)
    Loop
End With

' Replace double spaces with single
With doc.Content.Find
    .Text = "  "
    .Replacement.Text = " "
    Do While .Execute(Replace:=wdReplaceAll)
    Loop
End With

' Normalize paragraph formatting
For Each para In doc.Paragraphs
    With para.Format
        .PageBreakBefore = False
        .KeepWithNext = False
        .KeepTogether = False
        .LineSpacingRule = wdLineSpaceDouble
        .SpaceBefore = 0
        .SpaceAfter = 0
    End With
Next para

' Single column layout for all sections
Dim sec As Section
For Each sec In doc.Sections
    sec.PageSetup.TextColumns.SetCount NumColumns:=1
Next sec

End Sub
