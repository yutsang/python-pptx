' Ground truth for "how tall is this text, really".
'
' Everything else in this repo MODELS PowerPoint's line breaking (Pillow glyph
' widths, a 1.2x line pitch, a 3pt paragraph gap). Those numbers were
' calibrated against this measurement and have been wrong before -- most
' recently a commentary box that measured 101% full while visibly showing a
' large gap at the bottom. Only PowerPoint knows where it actually broke the
' lines, and this is how to ask it.
'
' HOW TO RUN
'   1. Open the exported .pptx in PowerPoint.
'   2. Alt+F11 (Developer > Visual Basic).
'   3. Insert > Module, paste this whole file in.
'   4. F5 (or Run > Run Sub). A message box reports every commentary box.
'   5. Copy the text out of the Immediate window (Ctrl+G) and paste it back.
'
' WHAT THE NUMBERS MEAN
'   BoundHeight  - the height PowerPoint's own layout engine gives the text
'   Lines        - how many lines it actually broke into
'   ShapeHeight  - the box, for comparison
'   Gap          - ShapeHeight - BoundHeight, i.e. the blank space you can see
'
' A large positive Gap next to a "fill=100%" report from inspect_pptx.py means
' this repo's model over-counts: it predicts more lines than PowerPoint draws.
' Compare Lines against the "used=NN.NL" figure to see by how much.

Option Explicit

Sub MeasureCommentaryBoxes()
    Dim sld As Slide
    Dim shp As Shape
    Dim msg As String
    Dim bh As Single, sh As Single, gap As Single
    Dim nLines As Long

    msg = "slide  shape                  Lines  BoundHeight  ShapeHeight   Gap" & vbCrLf
    msg = msg & String(74, "-") & vbCrLf

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame2.HasText Then
                    If InStr(1, shp.Name, "textMainBullets", vbTextCompare) = 1 _
                       Or InStr(1, shp.Name, "coSummaryShape", vbTextCompare) = 1 Then

                        bh = shp.TextFrame2.TextRange.BoundHeight
                        sh = shp.Height
                        gap = sh - bh
                        nLines = shp.TextFrame2.TextRange.Lines.Count

                        msg = msg & Format(sld.SlideIndex, "@@@@ ") & "  " & _
                              Left(shp.Name & String(22, " "), 22) & _
                              Format(nLines, "@@@@@") & "  " & _
                              Format(Round(bh, 1), "@@@@@@@@@@") & "  " & _
                              Format(Round(sh, 1), "@@@@@@@@@@@") & "  " & _
                              Format(Round(gap, 1), "@@@@@@") & vbCrLf

                        ' Also to the Immediate window (Ctrl+G) so it can be
                        ' copied as text rather than retyped from a dialog.
                        Debug.Print sld.SlideIndex; shp.Name; nLines; _
                                    Round(bh, 1); Round(sh, 1); Round(gap, 1)
                    End If
                End If
            End If
        Next shp
    Next sld

    MsgBox msg, vbOKOnly, "Rendered text height (PowerPoint's own numbers)"
End Sub
