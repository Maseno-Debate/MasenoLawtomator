' ================================
'   citation.bas
'   Footnote & Citation Formatter
'   Complements formatting.bas
'   Triggered via CTRL + ALT + A
' ================================

Option Explicit

' ----------------------------------------
'   MAIN WRAPPER (Footnotes only)
' ----------------------------------------
Sub FormatCitations()
    Call FormatFootnotes
End Sub

' ----------------------------------------
'   FORMAT FOOTNOTES:
'   Times New Roman, Size 10,
'   Slight Grey (wdColorGray50)
' ----------------------------------------
Sub FormatFootnotes()
    Dim fn As Footnote

    For Each fn In ActiveDocument.Footnotes
        ' Font settings
        With fn.Range.Font
            .Name = "Times New Roman"
            .Size = 10
            .Italic = False
            .Color = RGB(80, 80, 80)
        End With

        ' Paragraph formatting
        With fn.Range.ParagraphFormat
            .LineSpacingRule = wdLineSpaceSingle
            .Alignment = wdAlignParagraphLeft
            .FirstLineIndent = -18
            .LeftIndent = 18
            .SpaceBefore = 0
            .SpaceAfter = 0
        End With
    Next fn
End Sub