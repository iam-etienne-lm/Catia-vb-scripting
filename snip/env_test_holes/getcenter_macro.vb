Sub CATMain()
'Document part
Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim partDocument1 As PartDocument
Set partDocument1 = documents1.Item("Part1.CATPart")

Dim part1 As Part
Set part1 = partDocument1.Part

'HSF part
Dim hybridShapeFactory1 As HybridShapeFactory
Dim hybridBodies1 As HybridBodies
Dim hybridBody1 As HybridBody
Dim hybridShapes1 As HybridShapes
Dim hybridShapeCircleCtrRad1 As HybridShapeCircleCtrRad
Dim reference1 As Reference
Dim hybridShapePointCenter1 As HybridShapePointCenter

Set hybridShapeFactory1 = part1.HybridShapeFactory
Set hybridBodies1 = part1.HybridBodies
Set hybridBody1 = hybridBodies1.Item("Geometrical Set.1")
Set hybridShapes1 = hybridBody1.HybridShapes
Set hybridShapeCircleCtrRad1 = hybridShapes1.Item("Circle.2")
Set reference1 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)
Set hybridShapePointCenter1 = hybridShapeFactory1.AddNewPointCenter(reference1)

hybridBody1.AppendHybridShape hybridShapePointCenter1

part1.InWorkObject = hybridShapePointCenter1

'part1.Update

End Sub

