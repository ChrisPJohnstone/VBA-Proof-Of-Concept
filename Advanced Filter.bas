Sub Advanced Filter()
  
Application.DisplayAlerts = False

Dim RgInput As Range, RgFilter As Range, RgOutput As Range

Set RgInput = Range("A1").CurrentRegion
Set RgFilter = Range("D1").CurrentRegion
Set RgOutput = Range("G1").CurrentRegion

RgInput.AdvancedFilter xlFilterCopy, RgFilter, RgOutput

End Sub
