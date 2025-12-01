' ================================
'   formatting.bas  
'   Legal Document Formatter
' ================================

Option Explicit

' -------------------------------
'   MAIN WRAPPER FUNCTION
' -------------------------------
Sub FormatWholeDocument()
    Call SetMarginsLegal
    Call FormatParagraphFont
    Call FormatParagraphSpacing
    Call FormatParagraphAlignment

    MsgBox "Document formatting completed.", vbInformation
End Sub

' -------------------------------
'   SET STANDARD LEGAL MARGINS
' -------------------------------
Sub SetMarginsLegal()
    With ActiveDocument.PageSetup
        .TopMargin = CentimetersToPoints(2.54)
        .BottomMargin = CentimetersToPoints(2.54)
        .LeftMargin = CentimetersToPoints(2.54)
        .RightMargin = CentimetersToPoints(2.54)
    End With
End Sub

' -------------------------------
'   FONT: TNR, SIZE 12, BLACK
' -------------------------------
Sub FormatParagraphFont()
    Dim para As Paragraph

    For Each para In ActiveDocument.Paragraphs
        With para.Range.Font
            .Name = "Times New Roman"
            .Size = 12
            .Color = wdColorBlack
        End With
    Next para
End Sub

' -------------------------------
'   LINE SPACING: 1.5
' -------------------------------
Sub FormatParagraphSpacing()
    Dim para As Paragraph

    For Each para In ActiveDocument.Paragraphs
        para.Range.ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
    Next para
End Sub

' -------------------------------
'   JUSTIFY (SHORT LINES = LEFT)
' -------------------------------
Sub FormatParagraphAlignment()
    Dim para As Paragraph

    For Each para In ActiveDocument.Paragraphs
        para.Range.ParagraphFormat.Alignment = wdAlignParagraphJustify
    Next para
End Sub