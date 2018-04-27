Sub CATMain()
'simple y without sketch

Dim productDocument1 As ProductDocument
Set productDocument1 = CATIA.ActiveDocument

Dim product1 As Product
Set product1 = productDocument1.Product

product1.Update

Dim products1 As Products
Set products1 = product1.Products

Dim product2 As Product
Set product2 = products1.AddNewComponent("Part", "part_501")

Dim move1 As Move
Set move1 = product2.Move

Set move1 = move1.MovableObject

Dim arrayOfVariantOfDouble1(11)
arrayOfVariantOfDouble1(0) = 1#
arrayOfVariantOfDouble1(1) = 0#
arrayOfVariantOfDouble1(2) = 0#
arrayOfVariantOfDouble1(3) = 0#
arrayOfVariantOfDouble1(4) = 1#
arrayOfVariantOfDouble1(5) = 0#
arrayOfVariantOfDouble1(6) = 0#
arrayOfVariantOfDouble1(7) = 0#
arrayOfVariantOfDouble1(8) = 1#
arrayOfVariantOfDouble1(9) = 0#
arrayOfVariantOfDouble1(10) = 0#
arrayOfVariantOfDouble1(11) = 0#
Set move1Variant = move1
move1Variant.Apply arrayOfVariantOfDouble1

Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim partDocument1 As PartDocument
Set partDocument1 = documents1.Item("part_501.CATPart")

Dim part1 As Part
Set part1 = partDocument1.Part

Dim hybridShapeFactory1 As HybridShapeFactory
Set hybridShapeFactory1 = part1.HybridShapeFactory

Dim axisSystems1 As AxisSystems
Set axisSystems1 = part1.AxisSystems

Dim axisSystem1 As AxisSystem
Set axisSystem1 = axisSystems1.Item("Absolute Axis System")

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromBRepName("RSur:(Face:(Brp:(AxisSystem.1;1);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", axisSystem1)

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromBRepName("FVertex:(Vertex:(Neighbours:(Face:(Brp:(AxisSystem.1;2);None:();Cf11:());Face:(Brp:(AxisSystem.1;3);None:();Cf11:());Face:(Brp:(AxisSystem.1;1);None:();Cf11:()));Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", axisSystem1)

Dim hybridShapeLineNormal1 As HybridShapeLineNormal
Set hybridShapeLineNormal1 = hybridShapeFactory1.AddNewLineNormal(reference1, reference2, 0#, -20#, False)

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Geometrical Set.1")

hybridBody1.AppendHybridShape hybridShapeLineNormal1

part1.InWorkObject = hybridShapeLineNormal1

part1.Update

Dim hybridShapeDirection1 As HybridShapeDirection
Set hybridShapeDirection1 = hybridShapeFactory1.AddNewDirectionByCoord(1#, 2#, 3#)

Dim hybridBody2 As HybridBody
Set hybridBody2 = hybridBodies1.Item("External References")

Dim hybridShapes1 As HybridShapes
Set hybridShapes1 = hybridBody2.HybridShapes

Dim hybridShapeCurveExplicit1 As HybridShapeCurveExplicit
Set hybridShapeCurveExplicit1 = hybridShapes1.Item("Curve.1")

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)

Dim hybridShapeExtremum1 As HybridShapeExtremum
Set hybridShapeExtremum1 = hybridShapeFactory1.AddNewExtremum(reference3, hybridShapeDirection1, 1)

hybridBody1.AppendHybridShape hybridShapeExtremum1

part1.InWorkObject = hybridShapeExtremum1

part1.Update

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(hybridShapeCurveExplicit1)

Dim reference5 As Reference
Set reference5 = part1.CreateReferenceFromObject(hybridShapeExtremum1)

Dim reference6 As Reference
Set reference6 = part1.CreateReferenceFromBRepName("RSur:(Face:(Brp:(AxisSystem.1;3);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", axisSystem1)

Dim hybridShapeDirection2 As HybridShapeDirection
Set hybridShapeDirection2 = hybridShapeFactory1.AddNewDirection(reference6)

Dim hybridShapePointOnCurve1 As HybridShapePointOnCurve
Set hybridShapePointOnCurve1 = hybridShapeFactory1.AddNewPointOnCurveWithReferenceAlongDirection(reference4, reference5, 0#, False, hybridShapeDirection2)

hybridBody1.AppendHybridShape hybridShapePointOnCurve1

part1.InWorkObject = hybridShapePointOnCurve1

part1.Update

Dim reference7 As Reference
Set reference7 = part1.CreateReferenceFromObject(hybridShapePointOnCurve1)

Dim reference8 As Reference
Set reference8 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(GSMLine.1;1);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", hybridShapeLineNormal1)

Dim hybridShapeNear1 As HybridShapeNear
Set hybridShapeNear1 = hybridShapeFactory1.AddNewNear(reference7, reference8)

hybridBody1.AppendHybridShape hybridShapeNear1

part1.InWorkObject = hybridShapeNear1

part1.Update

Dim reference9 As Reference
Set reference9 = part1.CreateReferenceFromObject(hybridShapeNear1)

Dim reference10 As Reference
Set reference10 = part1.CreateReferenceFromBRepName("RSur:(Face:(Brp:(AxisSystem.1;3);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", axisSystem1)

Dim hybridShapeSymmetry1 As HybridShapeSymmetry
Set hybridShapeSymmetry1 = hybridShapeFactory1.AddNewSymmetry(reference9, reference10)

hybridShapeSymmetry1.VolumeResult = False

hybridBody1.AppendHybridShape hybridShapeSymmetry1

part1.InWorkObject = hybridShapeSymmetry1

part1.Update

Dim reference11 As Reference
Set reference11 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(GSMLine.1;2);None:(Limits1:();Limits2:();-1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", hybridShapeLineNormal1)

Dim reference12 As Reference
Set reference12 = part1.CreateReferenceFromObject(hybridShapeNear1)

Dim hybridShapeLinePtPt1 As HybridShapeLinePtPt
Set hybridShapeLinePtPt1 = hybridShapeFactory1.AddNewLinePtPtExtended(reference11, reference12, 0#, 2#)

hybridBody1.AppendHybridShape hybridShapeLinePtPt1

part1.InWorkObject = hybridShapeLinePtPt1

part1.Update

Dim reference13 As Reference
Set reference13 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

Dim reference14 As Reference
Set reference14 = part1.CreateReferenceFromBRepName("RSur:(Face:(Brp:(AxisSystem.1;3);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", axisSystem1)

Dim hybridShapeSymmetry2 As HybridShapeSymmetry
Set hybridShapeSymmetry2 = hybridShapeFactory1.AddNewSymmetry(reference13, reference14)

hybridShapeSymmetry2.VolumeResult = False

hybridBody1.AppendHybridShape hybridShapeSymmetry2

part1.InWorkObject = hybridShapeSymmetry2

part1.Update

Dim reference15 As Reference
Set reference15 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

Dim reference16 As Reference
Set reference16 = part1.CreateReferenceFromObject(hybridShapeNear1)

Dim hybridShapePlaneNormal1 As HybridShapePlaneNormal
Set hybridShapePlaneNormal1 = hybridShapeFactory1.AddNewPlaneNormal(reference15, reference16)

hybridBody1.AppendHybridShape hybridShapePlaneNormal1

part1.InWorkObject = hybridShapePlaneNormal1

part1.Update

Dim reference17 As Reference
Set reference17 = part1.CreateReferenceFromObject(hybridShapeNear1)

Dim reference18 As Reference
Set reference18 = part1.CreateReferenceFromObject(hybridShapePlaneNormal1)

Dim hybridShapeCircleCtrRad1 As HybridShapeCircleCtrRad
Set hybridShapeCircleCtrRad1 = hybridShapeFactory1.AddNewCircleCtrRadWithAngles(reference17, reference18, False, 0.5, 0#, 180#)

hybridShapeCircleCtrRad1.SetLimitation 0

hybridBody1.AppendHybridShape hybridShapeCircleCtrRad1

part1.InWorkObject = hybridShapeCircleCtrRad1

part1.Update

Dim reference19 As Reference
Set reference19 = part1.CreateReferenceFromObject(hybridShapeNear1)

Dim reference20 As Reference
Set reference20 = part1.CreateReferenceFromObject(hybridShapePlaneNormal1)

Set hybridShapeCircleCtrRad1 = hybridShapeFactory1.AddNewCircleCtrRad(reference19, reference20, False, 0.5)

hybridShapeCircleCtrRad1.SetLimitation 1

hybridBody1.AppendHybridShape hybridShapeCircleCtrRad1

part1.InWorkObject = hybridShapeCircleCtrRad1

part1.Update

Dim reference21 As Reference
Set reference21 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(GSMLine.1;2);None:(Limits1:();Limits2:();-1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", hybridShapeLineNormal1)

Dim reference22 As Reference
Set reference22 = part1.CreateReferenceFromBRepName("RSur:(Face:(Brp:(AxisSystem.1;1);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", axisSystem1)

Dim hybridShapeCircleCtrRad2 As HybridShapeCircleCtrRad
Set hybridShapeCircleCtrRad2 = hybridShapeFactory1.AddNewCircleCtrRad(reference21, reference22, False, 0.5)

hybridShapeCircleCtrRad2.SetLimitation 1

hybridBody1.AppendHybridShape hybridShapeCircleCtrRad2

part1.InWorkObject = hybridShapeCircleCtrRad2

part1.Update

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("PartBody")

part1.InWorkObject = body1

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim reference23 As Reference
Set reference23 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad1)

Dim reference24 As Reference
Set reference24 = part1.CreateReferenceFromObject(hybridShapeLinePtPt1)

Dim rib1 As Rib
Set rib1 = shapeFactory1.AddNewRibFromRef(reference23, reference24)

part1.UpdateObject rib1

part1.UpdateObject rib1

Dim reference25 As Reference
Set reference25 = part1.CreateReferenceFromObject(hybridShapeCircleCtrRad2)

Dim reference26 As Reference
Set reference26 = part1.CreateReferenceFromObject(hybridShapeLineNormal1)

Dim rib2 As Rib
Set rib2 = shapeFactory1.AddNewRibFromRef(reference25, reference26)

part1.UpdateObject rib2

Dim shapes1 As Shapes
Set shapes1 = body1.Shapes

Dim mirror1 As Mirror
Set mirror1 = shapes1.Item("Mirror.1")

Dim reference27 As Reference
Set reference27 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(AxisSystem.1;3);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithInitialFeatureSupport;MFBRepVersion_CXR15)", axisSystem1)

mirror1.MirroringPlane = reference27

part1.UpdateObject mirror1

Dim mirror2 As Mirror
Set mirror2 = shapes1.Item("Mirror.2")

Dim reference28 As Reference
Set reference28 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(AxisSystem.1;3);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithInitialFeatureSupport;MFBRepVersion_CXR15)", axisSystem1)

mirror2.MirroringPlane = reference28

part1.UpdateObject mirror2

Dim reference29 As Reference
Set reference29 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(AxisSystem.1;3);None:();Cf11:());WithTemporaryBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", axisSystem1)

Dim mirror3 As Mirror
Set mirror3 = shapeFactory1.AddNewMirror(reference29)

part1.UpdateObject mirror3

Dim selection1 As Selection
Set selection1 = productDocument1.Selection

Dim visPropertySet1 As VisPropertySet
Set visPropertySet1 = selection1.VisProperties

Dim hybridShapes2 As HybridShapes
Set hybridShapes2 = hybridShapeSymmetry2.Parent

Dim bSTR1 As String
bSTR1 = hybridShapeSymmetry2.Name

selection1.Add hybridShapeSymmetry2

Set visPropertySet1 = visPropertySet1.Parent

Dim bSTR2 As String
bSTR2 = visPropertySet1.Name

Dim bSTR3 As String
bSTR3 = visPropertySet1.Name

visPropertySet1.SetShow 1

selection1.Clear

End Sub

