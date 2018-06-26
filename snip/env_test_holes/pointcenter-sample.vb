Option Explicit

Sub CATMain()
'Document part
Dim productDocument1 As ProductDocument
Set productDocument1 = CATIA.ActiveDocument

Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim partDocument1 As PartDocument
Set partDocument1 = documents1.Item(2)

Dim part1 As Part
Dim opart As Part
Set opart = partDocument1.Part

'Additional dim
Dim reference1 As Reference
Dim InputObject(0)
Dim oCentre 'As Selection
Dim status
Dim oTemp 'As Item
Dim oRef As Reference
Dim oHb
Dim oPoint

'HSF part
Dim oHsf As HybridShapeFactory
Set oHsf = opart.HybridShapeFactory


'Selection part
InputObject(0) = "Edge"
Set oCentre = CATIA.ActiveDocument.Selection
status = oCentre.SelectElement2(InputObject, "Select Circle", False)

Set oTemp = oCentre.Item(1)
Set oRef = oTemp.Reference 'oref = oCentre.Item(1).Reference
Set oPoint = oHsf.AddNewPointCenter(oRef)
Set oHb = opart.HybridBodies.Add()
oPoint.Compute
oHb.AppendHybridShape oPoint

End Sub


