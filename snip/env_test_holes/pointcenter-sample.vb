Sub CATMain()

Dim productDocument1 As ProductDocument
Set productDocument1 = CATIA.ActiveDocument
Dim documents1 As Documents
Set documents1 = CATIA.Documents
Dim partDocument1 As PartDocument
Set partDocument1 = documents1.Item(2)
Dim part1 As Part
Set oPart = partDocument1.Part
Dim reference1 As Reference
Set oHSF = oPart.HybridShapeFactory
Dim InputObject(0)

InputObject(0) = "Edge"
Set oCentre = CATIA.ActiveDocument.Selection
Status = oCentre.SelectElement2(InputObject, "Select Circle", False)
Set oTemp = oCentre.Item(1)
Set oRef = oTemp.Reference
Set oPoint = oHSF.AddNewPointCenter(oRef)
Set oHB = oPart.HybridBodies.Add()
oPoint.Compute
oHB.AppendHybridShape oPoint

End Sub
