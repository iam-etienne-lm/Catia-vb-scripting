Sub CATMain()
'similar method as other snip, instead here pts are created


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
Set reference1 = part1.CreateReferenceFromBRepName("RSur:(Face:(Brp:(Pad.3;3:(Brp:(Sketch.2;5)));None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", pad1)

Dim hybridShapeExtract1 As HybridShapeExtract
Set hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference1)

hybridShapeExtract1.PropagationType = 3

hybridShapeExtract1.ComplementaryExtract = False

hybridShapeExtract1.IsFederated = False

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Geometrical Set.1")

hybridBody1.AppendHybridShape hybridShapeExtract1

part1.InWorkObject = hybridShapeExtract1

part1.Update

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromBRepName("BorderREdge:(BEdge:(Brp:(FeatureRSUR.8;(Brp:(Pad.3;3:(Brp:(Sketch.2;5)));Brp:(Pad.3;2)));None:(Limits1:();Limits2:();-1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", hybridShapeExtract1)

Dim hybridShapePointCenter1 As HybridShapePointCenter
Set hybridShapePointCenter1 = hybridShapeFactory1.AddNewPointCenter(reference2)

hybridBody1.AppendHybridShape hybridShapePointCenter1

part1.InWorkObject = hybridShapePointCenter1

part1.Update

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromBRepName("BorderREdge:(BEdge:(Brp:(FeatureRSUR.8;(Brp:(Pad.3;3:(Brp:(Sketch.2;5)));Brp:(Pad.3;1)));None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", hybridShapeExtract1)

Dim hybridShapePointCenter2 As HybridShapePointCenter
Set hybridShapePointCenter2 = hybridShapeFactory1.AddNewPointCenter(reference3)

hybridBody1.AppendHybridShape hybridShapePointCenter2

part1.InWorkObject = hybridShapePointCenter2

part1.Update

End Sub

