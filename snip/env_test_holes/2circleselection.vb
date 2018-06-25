Sub CATMain()

  'selection from simple circle  
Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim partDocument1 As PartDocument
Set partDocument1 = documents1.Item("Part4.CATPart")

Dim part1 As Part
Set part1 = partDocument1.Part

Dim hybridShapeFactory1 As HybridShapeFactory
Set hybridShapeFactory1 = part1.HybridShapeFactory

Dim hybridShapeDirection1 As HybridShapeDirection
Set hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1#, 2#, 3#)

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Geometrical Set.1")

Dim hybridShapes1 As HybridShapes
Set hybridShapes1 = hybridBody1.HybridShapes

Dim hybridShapeCircleCtrRad1 As HybridShapeCircleCtrRad
Set hybridShapeCircleCtrRad1 = hybridShapes1.Item("Circle.1")

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)

Dim hybridShapeExtremum1 As HybridShapeExtremum
Set hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference1, hybridShapeDirection1, 1)

hybridBody1.AppendHybridShape hybridShapeExtremum1

part1.InWorkObject = hybridShapeExtremum1

part1.Update

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)

Dim hybridShapePointCenter1 As HybridShapePointCenter
Set hybridShapePointCenter1 = hybridShapeFactory1.AddNewPointCenter(reference2)

hybridBody1.AppendHybridShape hybridShapePointCenter1

part1.InWorkObject = hybridShapePointCenter1

part1.Update

End Sub
