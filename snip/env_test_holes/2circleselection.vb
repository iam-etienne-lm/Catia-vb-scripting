Sub CATMain()

'Selection from pad limit. too complex
Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim partDocument1 As PartDocument
Set partDocument1 = documents1.Item("Holes 8 to 20.CATPart")

Dim part1 As Part
Set part1 = partDocument1.Part

Dim hybridShapeFactory1 As HybridShapeFactory
Set hybridShapeFactory1 = part1.HybridShapeFactory

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("PartBody")

Dim shapes1 As Shapes
Set shapes1 = body1.Shapes

Dim pad1 As Pad
Set pad1 = shapes1.Item("Pad.3")

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromBRepName("REdge:(Edge:(Face:(Brp:(Pad.3;2);None:();Cf11:());Face:(Brp:(Pad.3;3:(Brp:(Sketch.2;3)));None:();Cf11:());None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", pad1)

Dim hybridShapePointCenter1 As HybridShapePointCenter
Set hybridShapePointCenter1 = hybridShapeFactory1.AddNewPointCenter(reference1)

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Geometrical Set.1")

hybridBody1.AppendHybridShape hybridShapePointCenter1

part1.InWorkObject = hybridShapePointCenter1

part1.Update

End Sub

