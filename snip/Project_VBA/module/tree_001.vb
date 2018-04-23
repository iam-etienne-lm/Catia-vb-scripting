Sub CATMain()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As Part
Set part1 = partDocument1.Part

Dim bodies1 As Bodies
Set bodies1 = part1.Bodies

Dim body1 As Body
Set body1 = bodies1.Item("PartBody")

Dim sketches1 As Sketches
Set sketches1 = body1.Sketches

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromName("Selection_RSur:(Face:(Brp:(AxisSystem.1;2);None:();Cf11:());AxisSystem.1;Z0;G4252)")

Dim sketch1 As Sketch
Set sketch1 = sketches1.Add(reference1)

Dim arrayOfVariantOfDouble1(8)
arrayOfVariantOfDouble1(0) = 0#
arrayOfVariantOfDouble1(1) = 0#
arrayOfVariantOfDouble1(2) = 0#
arrayOfVariantOfDouble1(3) = 0#
arrayOfVariantOfDouble1(4) = 1#
arrayOfVariantOfDouble1(5) = 0#
arrayOfVariantOfDouble1(6) = 0#
arrayOfVariantOfDouble1(7) = 0#
arrayOfVariantOfDouble1(8) = 1#
Set sketch1Variant = sketch1
sketch1Variant.SetAbsoluteAxisData arrayOfVariantOfDouble1

part1.InWorkObject = sketch1

Dim factory2D1 As Factory2D
Set factory2D1 = sketch1.OpenEdition()

Dim geometricElements1 As GeometricElements
Set geometricElements1 = sketch1.GeometricElements

Dim axis2D1 As Axis2D
Set axis2D1 = geometricElements1.Item("AbsoluteAxis")

Dim line2D1 As Line2D
Set line2D1 = axis2D1.GetItem("HDirection")

line2D1.ReportName = 1

Dim line2D2 As Line2D
Set line2D2 = axis2D1.GetItem("VDirection")

line2D2.ReportName = 2

Dim circle2D1 As Circle2D
Set circle2D1 = factory2D1.CreateClosedCircle(0#, 0#, 5#)

Dim point2D1 As Point2D
Set point2D1 = axis2D1.GetItem("Origin")

circle2D1.CenterPoint = point2D1

circle2D1.ReportName = 3

circle2D1.Construction = True

Dim constraints1 As Constraints
Set constraints1 = sketch1.Constraints

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromObject(circle2D1)

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddMonoEltCst(catCstTypeRadius, reference2)

constraint1.Mode = catCstModeDrivingDimension

Dim length1 As Length
Set length1 = constraint1.Dimension

length1.Value = 5#

Dim point2D2 As Point2D
Set point2D2 = factory2D1.CreatePoint(1.394122, 6.117169)

point2D2.ReportName = 4

Dim point2D3 As Point2D
Set point2D3 = factory2D1.CreatePoint(6.906823, -0.452613)

point2D3.ReportName = 5

Dim line2D3 As Line2D
Set line2D3 = factory2D1.CreateLine(1.394122, 6.117169, 6.906823, -0.452613)

line2D3.ReportName = 6

line2D3.Construction = True

line2D3.StartPoint = point2D2

line2D3.EndPoint = point2D3

Dim point2D4 As Point2D
Set point2D4 = factory2D1.CreatePoint(-5.439608, 1.295946)

point2D4.ReportName = 7

Dim point2D5 As Point2D
Set point2D5 = factory2D1.CreatePoint(-1.731486, 5.715115)

point2D5.ReportName = 8

Dim line2D4 As Line2D
Set line2D4 = factory2D1.CreateLine(-5.439608, 1.295946, -1.731486, 5.715115)

line2D4.ReportName = 9

line2D4.Construction = True

line2D4.StartPoint = point2D4

line2D4.EndPoint = point2D5

Dim axisSystems1 As AxisSystems
Set axisSystems1 = part1.AxisSystems

Dim axisSystem1 As AxisSystem
Set axisSystem1 = axisSystems1.Item("Absolute Axis System")

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(AxisSystem.1;3);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", axisSystem1)

Dim geometricElements2 As GeometricElements
Set geometricElements2 = factory2D1.CreateIntersections(reference3)

Dim geometry2D1 As Geometry2D
Set geometry2D1 = geometricElements2.Item("Mark.1")

geometry2D1.Construction = True

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(geometry2D1)

Dim reference5 As Reference
Set reference5 = part1.CreateReferenceFromObject(line2D4)

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddBiEltCst(catCstTypeAngle, reference4, reference5)

constraint2.Mode = catCstModeDrivingDimension

constraint2.AngleSector = catCstAngleSector3

Dim reference6 As Reference
Set reference6 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(AxisSystem.1;3);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", axisSystem1)

Dim geometricElements3 As GeometricElements
Set geometricElements3 = factory2D1.CreateIntersections(reference6)

Dim geometry2D2 As Geometry2D
Set geometry2D2 = geometricElements3.Item("Mark.1")

geometry2D2.Construction = True

Dim reference7 As Reference
Set reference7 = part1.CreateReferenceFromObject(line2D3)

Dim reference8 As Reference
Set reference8 = part1.CreateReferenceFromObject(geometry2D2)

Dim constraint3 As Constraint
Set constraint3 = constraints1.AddBiEltCst(catCstTypeAngle, reference7, reference8)

constraint3.Mode = catCstModeDrivingDimension

constraint3.AngleSector = catCstAngleSector2

Dim reference9 As Reference
Set reference9 = part1.CreateReferenceFromObject(line2D3)

Dim reference10 As Reference
Set reference10 = part1.CreateReferenceFromObject(circle2D1)

Dim constraint4 As Constraint
Set constraint4 = constraints1.AddBiEltCst(catCstTypeTangency, reference9, reference10)

constraint4.Mode = catCstModeDrivingDimension

Dim reference11 As Reference
Set reference11 = part1.CreateReferenceFromObject(line2D4)

Dim reference12 As Reference
Set reference12 = part1.CreateReferenceFromObject(circle2D1)

Dim constraint5 As Constraint
Set constraint5 = constraints1.AddBiEltCst(catCstTypeTangency, reference11, reference12)

constraint5.Mode = catCstModeDrivingDimension

Dim point2D6 As Point2D
Set point2D6 = factory2D1.CreatePoint(3.830222, 3.213938)

point2D6.ReportName = 10

Dim reference13 As Reference
Set reference13 = part1.CreateReferenceFromObject(line2D3)

Dim reference14 As Reference
Set reference14 = part1.CreateReferenceFromObject(point2D6)

Dim constraint6 As Constraint
Set constraint6 = constraints1.AddBiEltCst(catCstTypeOn, reference13, reference14)

constraint6.Mode = catCstModeDrivingDimension

Dim reference15 As Reference
Set reference15 = part1.CreateReferenceFromObject(circle2D1)

Dim reference16 As Reference
Set reference16 = part1.CreateReferenceFromObject(point2D6)

Dim constraint7 As Constraint
Set constraint7 = constraints1.AddBiEltCst(catCstTypeOn, reference15, reference16)

constraint7.Mode = catCstModeDrivingDimension

Dim point2D7 As Point2D
Set point2D7 = factory2D1.CreatePoint(-3.830222, 3.213938)

point2D7.ReportName = 11

Dim reference17 As Reference
Set reference17 = part1.CreateReferenceFromObject(circle2D1)

Dim reference18 As Reference
Set reference18 = part1.CreateReferenceFromObject(point2D7)

Dim constraint8 As Constraint
Set constraint8 = constraints1.AddBiEltCst(catCstTypeOn, reference17, reference18)

constraint8.Mode = catCstModeDrivingDimension

Dim reference19 As Reference
Set reference19 = part1.CreateReferenceFromObject(line2D4)

Dim reference20 As Reference
Set reference20 = part1.CreateReferenceFromObject(point2D7)

Dim constraint9 As Constraint
Set constraint9 = constraints1.AddBiEltCst(catCstTypeOn, reference19, reference20)

constraint9.Mode = catCstModeDrivingDimension

Dim point2D8 As Point2D
Set point2D8 = factory2D1.CreatePoint(-4.482926, 3.9918)

point2D8.ReportName = 12

Dim point2D9 As Point2D
Set point2D9 = factory2D1.CreatePoint(0#, -1.350743)

point2D9.ReportName = 13

Dim line2D5 As Line2D
Set line2D5 = factory2D1.CreateLine(-4.482926, 3.9918, -0#, -1.350743)

line2D5.ReportName = 14

line2D5.StartPoint = point2D8

line2D5.EndPoint = point2D9

Dim reference21 As Reference
Set reference21 = part1.CreateReferenceFromObject(point2D9)

Dim reference22 As Reference
Set reference22 = part1.CreateReferenceFromObject(line2D2)

Dim constraint10 As Constraint
Set constraint10 = constraints1.AddBiEltCst(catCstTypeOn, reference21, reference22)

constraint10.Mode = catCstModeDrivingDimension

Dim reference23 As Reference
Set reference23 = part1.CreateReferenceFromBRepName("FSur:(Face:(Brp:(AxisSystem.1;3);None:();Cf11:());WithPermanentBody;WithoutBuildError;WithInitialFeatureSupport;MonoFond;MFBRepVersion_CXR15)", axisSystem1)

Dim geometricElements4 As GeometricElements
Set geometricElements4 = factory2D1.CreateIntersections(reference23)

Dim geometry2D3 As Geometry2D
Set geometry2D3 = geometricElements4.Item("Mark.1")

geometry2D3.Construction = True

Dim reference24 As Reference
Set reference24 = part1.CreateReferenceFromObject(line2D5)

Dim reference25 As Reference
Set reference25 = part1.CreateReferenceFromObject(geometry2D3)

Dim constraint11 As Constraint
Set constraint11 = constraints1.AddBiEltCst(catCstTypeAngle, reference24, reference25)

constraint11.Mode = catCstModeDrivingDimension

constraint11.AngleSector = catCstAngleSector2

Dim point2D10 As Point2D
Set point2D10 = factory2D1.CreatePoint(0#, -6#)

point2D10.ReportName = 15

Dim line2D6 As Line2D
Set line2D6 = factory2D1.CreateLine(0#, -1.350743, 0#, -6#)

line2D6.ReportName = 16

line2D6.StartPoint = point2D9

line2D6.EndPoint = point2D10

Dim reference26 As Reference
Set reference26 = part1.CreateReferenceFromObject(line2D6)

Dim reference27 As Reference
Set reference27 = part1.CreateReferenceFromObject(line2D2)

Dim constraint12 As Constraint
Set constraint12 = constraints1.AddBiEltCst(catCstTypeVerticality, reference26, reference27)

constraint12.Mode = catCstModeDrivingDimension

Dim reference28 As Reference
Set reference28 = part1.CreateReferenceFromObject(point2D7)

Dim reference29 As Reference
Set reference29 = part1.CreateReferenceFromObject(line2D5)

Dim constraint13 As Constraint
Set constraint13 = constraints1.AddBiEltCst(catCstTypeOn, reference28, reference29)

constraint13.Mode = catCstModeDrivingDimension

Dim reference30 As Reference
Set reference30 = part1.CreateReferenceFromObject(point2D8)

Dim reference31 As Reference
Set reference31 = part1.CreateReferenceFromObject(line2D4)

Dim constraint14 As Constraint
Set constraint14 = constraints1.AddBiEltCst(catCstTypeDistance, reference30, reference31)

constraint14.Mode = catCstModeDrivingDimension

Dim length2 As Length
Set length2 = constraint14.Dimension

length2.Value = 1#

Dim reference32 As Reference
Set reference32 = part1.CreateReferenceFromObject(point2D10)

Dim reference33 As Reference
Set reference33 = part1.CreateReferenceFromObject(circle2D1)

Dim constraint15 As Constraint
Set constraint15 = constraints1.AddBiEltCst(catCstTypeDistance, reference32, reference33)

constraint15.Mode = catCstModeDrivingDimension

Dim length3 As Length
Set length3 = constraint15.Dimension

length3.Value = 1#

Dim point2D11 As Point2D
Set point2D11 = factory2D1.CreatePoint(-2.830222, 1.141111)

point2D11.ReportName = 17

Dim point2D12 As Point2D
Set point2D12 = factory2D1.CreatePoint(-2.830222, 7.556188)

point2D12.ReportName = 18

Dim line2D7 As Line2D
Set line2D7 = factory2D1.CreateLine(-2.830222, 1.141111, -2.830222, 7.556188)

line2D7.ReportName = 19

line2D7.Construction = True

line2D7.StartPoint = point2D11

line2D7.EndPoint = point2D12

Dim reference34 As Reference
Set reference34 = part1.CreateReferenceFromObject(line2D7)

Dim reference35 As Reference
Set reference35 = part1.CreateReferenceFromObject(line2D2)

Dim constraint16 As Constraint
Set constraint16 = constraints1.AddBiEltCst(catCstTypeVerticality, reference34, reference35)

constraint16.Mode = catCstModeDrivingDimension

Dim reference36 As Reference
Set reference36 = part1.CreateReferenceFromObject(point2D7)

Dim reference37 As Reference
Set reference37 = part1.CreateReferenceFromObject(line2D7)

Dim constraint17 As Constraint
Set constraint17 = constraints1.AddBiEltCst(catCstTypeDistance, reference36, reference37)

constraint17.Mode = catCstModeDrivingDimension

Dim length4 As Length
Set length4 = constraint17.Dimension

length4.Value = 1#

Dim point2D13 As Point2D
Set point2D13 = factory2D1.CreatePoint(-2.830222, 5.121874)

point2D13.ReportName = 20

Dim point2D14 As Point2D
Set point2D14 = factory2D1.CreatePoint(-2.830222, 2.022184)

point2D14.ReportName = 21

Dim line2D8 As Line2D
Set line2D8 = factory2D1.CreateLine(-2.830222, 5.121874, -2.830222, 2.022184)

line2D8.ReportName = 22

line2D8.StartPoint = point2D13

line2D8.EndPoint = point2D14

Dim reference38 As Reference
Set reference38 = part1.CreateReferenceFromObject(point2D13)

Dim reference39 As Reference
Set reference39 = part1.CreateReferenceFromObject(line2D7)

Dim constraint18 As Constraint
Set constraint18 = constraints1.AddBiEltCst(catCstTypeOn, reference38, reference39)

constraint18.Mode = catCstModeDrivingDimension

Dim reference40 As Reference
Set reference40 = part1.CreateReferenceFromObject(point2D14)

Dim reference41 As Reference
Set reference41 = part1.CreateReferenceFromObject(line2D5)

Dim constraint19 As Constraint
Set constraint19 = constraints1.AddBiEltCst(catCstTypeOn, reference40, reference41)

constraint19.Mode = catCstModeDrivingDimension

Dim reference42 As Reference
Set reference42 = part1.CreateReferenceFromObject(line2D8)

Dim reference43 As Reference
Set reference43 = part1.CreateReferenceFromObject(line2D2)

Dim constraint20 As Constraint
Set constraint20 = constraints1.AddBiEltCst(catCstTypeVerticality, reference42, reference43)

constraint20.Mode = catCstModeDrivingDimension

Dim point2D15 As Point2D
Set point2D15 = factory2D1.CreatePoint(-2.830222, 4.121874)

point2D15.ReportName = 23

point2D15.Construction = True

Dim reference44 As Reference
Set reference44 = part1.CreateReferenceFromObject(circle2D1)

Dim reference45 As Reference
Set reference45 = part1.CreateReferenceFromObject(point2D15)

Dim constraint21 As Constraint
Set constraint21 = constraints1.AddBiEltCst(catCstTypeOn, reference44, reference45)

constraint21.Mode = catCstModeDrivingDimension

Dim reference46 As Reference
Set reference46 = part1.CreateReferenceFromObject(line2D8)

Dim reference47 As Reference
Set reference47 = part1.CreateReferenceFromObject(point2D15)

Dim constraint22 As Constraint
Set constraint22 = constraints1.AddBiEltCst(catCstTypeOn, reference46, reference47)

constraint22.Mode = catCstModeDrivingDimension

Dim point2D16 As Point2D
Set point2D16 = factory2D1.CreatePoint(-2.830222, -4.121874)

point2D16.ReportName = 24

point2D16.Construction = True

Dim reference48 As Reference
Set reference48 = part1.CreateReferenceFromObject(circle2D1)

Dim reference49 As Reference
Set reference49 = part1.CreateReferenceFromObject(point2D16)

Dim constraint23 As Constraint
Set constraint23 = constraints1.AddBiEltCst(catCstTypeOn, reference48, reference49)

constraint23.Mode = catCstModeDrivingDimension

Dim reference50 As Reference
Set reference50 = part1.CreateReferenceFromObject(line2D8)

Dim reference51 As Reference
Set reference51 = part1.CreateReferenceFromObject(point2D16)

Dim constraint24 As Constraint
Set constraint24 = constraints1.AddBiEltCst(catCstTypeOn, reference50, reference51)

constraint24.Mode = catCstModeDrivingDimension

Dim reference52 As Reference
Set reference52 = part1.CreateReferenceFromObject(point2D13)

Dim reference53 As Reference
Set reference53 = part1.CreateReferenceFromObject(point2D15)

Dim constraint25 As Constraint
Set constraint25 = constraints1.AddBiEltCst(catCstTypeDistance, reference52, reference53)

constraint25.Mode = catCstModeDrivingDimension

Dim length5 As Length
Set length5 = constraint25.Dimension

length5.Value = 1#

sketch1.CloseEdition

part1.InWorkObject = body1

part1.Update

length2.Value = 1#

length3.Value = 1#

length4.Value = 1#

length5.Value = 1#

length2.Value = 1#

length3.Value = 1#

length4.Value = 1#

length5.Value = 1#

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Add()

Dim hybridShapeFactory1 As HybridShapeFactory
Set hybridShapeFactory1 = part1.HybridShapeFactory

Dim reference54 As Reference
Set reference54 = part1.CreateReferenceFromBRepName("WireREdge:(Wire:(Brp:(Sketch.1;16);None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

Dim hybridShapeExtract1 As HybridShapeExtract
Set hybridShapeExtract1 = hybridShapeFactory1.AddNewExtract(reference54)

hybridShapeExtract1.PropagationType = 3

hybridShapeExtract1.ComplementaryExtract = False

hybridShapeExtract1.IsFederated = False

hybridBody1.AppendHybridShape hybridShapeExtract1

part1.InWorkObject = hybridShapeExtract1

part1.Update

Dim reference55 As Reference
Set reference55 = part1.CreateReferenceFromObject(hybridShapeExtract1)

Dim reference56 As Reference
Set reference56 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;15);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

Dim hybridShapePlaneNormal1 As HybridShapePlaneNormal
Set hybridShapePlaneNormal1 = hybridShapeFactory1.AddNewPlaneNormal(reference55, reference56)

hybridBody1.AppendHybridShape hybridShapePlaneNormal1

part1.InWorkObject = hybridShapePlaneNormal1

part1.Update

Dim sketches2 As Sketches
Set sketches2 = hybridBody1.HybridSketches

Dim hybridShapes1 As HybridShapes
Set hybridShapes1 = hybridBody1.HybridShapes

Dim reference57 As Reference
Set reference57 = hybridShapes1.Item("Plane.1")

Dim sketch2 As Sketch
Set sketch2 = sketches2.Add(reference57)

Dim arrayOfVariantOfDouble2(8)
arrayOfVariantOfDouble2(0) = 0#
arrayOfVariantOfDouble2(1) = 0#
arrayOfVariantOfDouble2(2) = -6#
arrayOfVariantOfDouble2(3) = 1#
arrayOfVariantOfDouble2(4) = 0#
arrayOfVariantOfDouble2(5) = 0#
arrayOfVariantOfDouble2(6) = -0#
arrayOfVariantOfDouble2(7) = 1#
arrayOfVariantOfDouble2(8) = 0#
Set sketch2Variant = sketch2
sketch2Variant.SetAbsoluteAxisData arrayOfVariantOfDouble2

part1.InWorkObject = sketch2

Dim factory2D2 As Factory2D
Set factory2D2 = sketch2.OpenEdition()

Dim geometricElements5 As GeometricElements
Set geometricElements5 = sketch2.GeometricElements

Dim axis2D2 As Axis2D
Set axis2D2 = geometricElements5.Item("AbsoluteAxis")

Dim line2D9 As Line2D
Set line2D9 = axis2D2.GetItem("HDirection")

line2D9.ReportName = 1

Dim line2D10 As Line2D
Set line2D10 = axis2D2.GetItem("VDirection")

line2D10.ReportName = 2

Dim circle2D2 As Circle2D
Set circle2D2 = factory2D2.CreateClosedCircle(0#, 0#, 0.4)

Dim point2D17 As Point2D
Set point2D17 = axis2D2.GetItem("Origin")

circle2D2.CenterPoint = point2D17

circle2D2.ReportName = 3

Dim constraints2 As Constraints
Set constraints2 = sketch2.Constraints

Dim reference58 As Reference
Set reference58 = part1.CreateReferenceFromObject(circle2D2)

Dim constraint26 As Constraint
Set constraint26 = constraints2.AddMonoEltCst(catCstTypeRadius, reference58)

constraint26.Mode = catCstModeDrivingDimension

sketch2.CloseEdition

part1.InWorkObject = hybridBody1

part1.Update

Dim reference59 As Reference
Set reference59 = part1.CreateReferenceFromBRepName("WireREdge:(Wire:(Brp:(Sketch.1;14);None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

Dim hybridShapeExtract2 As HybridShapeExtract
Set hybridShapeExtract2 = hybridShapeFactory1.AddNewExtract(reference59)

hybridShapeExtract2.PropagationType = 3

hybridShapeExtract2.ComplementaryExtract = False

hybridShapeExtract2.IsFederated = False

hybridBody1.AppendHybridShape hybridShapeExtract2

part1.InWorkObject = hybridShapeExtract2

part1.Update

Dim reference60 As Reference
Set reference60 = part1.CreateReferenceFromObject(hybridShapeExtract2)

Dim hybridShapePlaneNormal2 As HybridShapePlaneNormal
Set hybridShapePlaneNormal2 = hybridShapeFactory1.AddNewPlaneNormal(reference60, Nothing)

hybridBody1.AppendHybridShape hybridShapePlaneNormal2

part1.InWorkObject = hybridShapePlaneNormal2

part1.Update

Dim reference61 As Reference
Set reference61 = hybridShapes1.Item("Plane.2")

Dim sketch3 As Sketch
Set sketch3 = sketches2.Add(reference61)

Dim arrayOfVariantOfDouble3(8)
arrayOfVariantOfDouble3(0) = 0#
arrayOfVariantOfDouble3(1) = -1.576352
arrayOfVariantOfDouble3(2) = 1.878623
arrayOfVariantOfDouble3(3) = 1#
arrayOfVariantOfDouble3(4) = 0#
arrayOfVariantOfDouble3(5) = 0#
arrayOfVariantOfDouble3(6) = -0#
arrayOfVariantOfDouble3(7) = 0.766044
arrayOfVariantOfDouble3(8) = 0.642788
Set sketch3Variant = sketch3
sketch3Variant.SetAbsoluteAxisData arrayOfVariantOfDouble3

part1.InWorkObject = sketch3

Dim factory2D3 As Factory2D
Set factory2D3 = sketch3.OpenEdition()

Dim geometricElements6 As GeometricElements
Set geometricElements6 = sketch3.GeometricElements

Dim axis2D3 As Axis2D
Set axis2D3 = geometricElements6.Item("AbsoluteAxis")

Dim line2D11 As Line2D
Set line2D11 = axis2D3.GetItem("HDirection")

line2D11.ReportName = 1

Dim line2D12 As Line2D
Set line2D12 = axis2D3.GetItem("VDirection")

line2D12.ReportName = 2

Dim reference62 As Reference
Set reference62 = part1.CreateReferenceFromBRepName("WireREdge:(Wire:(Brp:(Sketch.1;14);None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MonoFond;MFBRepVersion_CXR15)", sketch1)

Dim geometricElements7 As GeometricElements
Set geometricElements7 = factory2D3.CreateIntersections(reference62)

Dim geometry2D4 As Geometry2D
Set geometry2D4 = geometricElements7.Item("Mark.1")

geometry2D4.Construction = True

Dim circle2D3 As Circle2D
Set circle2D3 = factory2D3.CreateClosedCircle(0#, -0.868241, 0.4)

circle2D3.ReportName = 3

Dim constraints3 As Constraints
Set constraints3 = sketch3.Constraints

Dim reference63 As Reference
Set reference63 = part1.CreateReferenceFromObject(circle2D3)

Dim constraint27 As Constraint
Set constraint27 = constraints3.AddMonoEltCst(catCstTypeRadius, reference63)

constraint27.Mode = catCstModeDrivingDimension

sketch3.CloseEdition

part1.InWorkObject = hybridBody1

part1.Update

Dim reference64 As Reference
Set reference64 = part1.CreateReferenceFromBRepName("WireREdge:(Wire:(Brp:(Sketch.1;22);None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

Dim hybridShapeExtract3 As HybridShapeExtract
Set hybridShapeExtract3 = hybridShapeFactory1.AddNewExtract(reference64)

hybridShapeExtract3.PropagationType = 3

hybridShapeExtract3.ComplementaryExtract = False

hybridShapeExtract3.IsFederated = False

hybridBody1.AppendHybridShape hybridShapeExtract3

part1.InWorkObject = hybridShapeExtract3

part1.Update

Dim reference65 As Reference
Set reference65 = part1.CreateReferenceFromObject(hybridShapeExtract3)

Dim reference66 As Reference
Set reference66 = part1.CreateReferenceFromBRepName("BorderFVertex:(BEdge:(Brp:(Sketch.1;20);None:(Limits1:();Limits2:();+1);Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

Dim hybridShapePlaneNormal3 As HybridShapePlaneNormal
Set hybridShapePlaneNormal3 = hybridShapeFactory1.AddNewPlaneNormal(reference65, reference66)

hybridBody1.AppendHybridShape hybridShapePlaneNormal3

part1.InWorkObject = hybridShapePlaneNormal3

part1.Update

Dim reference67 As Reference
Set reference67 = hybridShapes1.Item("Plane.3")

Dim sketch4 As Sketch
Set sketch4 = sketches2.Add(reference67)

Dim arrayOfVariantOfDouble4(8)
arrayOfVariantOfDouble4(0) = 0#
arrayOfVariantOfDouble4(1) = -0#
arrayOfVariantOfDouble4(2) = 5.121874
arrayOfVariantOfDouble4(3) = 1#
arrayOfVariantOfDouble4(4) = 0#
arrayOfVariantOfDouble4(5) = 0#
arrayOfVariantOfDouble4(6) = 0#
arrayOfVariantOfDouble4(7) = -1#
arrayOfVariantOfDouble4(8) = -0#
Set sketch4Variant = sketch4
sketch4Variant.SetAbsoluteAxisData arrayOfVariantOfDouble4

part1.InWorkObject = sketch4

Dim factory2D4 As Factory2D
Set factory2D4 = sketch4.OpenEdition()

Dim geometricElements8 As GeometricElements
Set geometricElements8 = sketch4.GeometricElements

Dim axis2D4 As Axis2D
Set axis2D4 = geometricElements8.Item("AbsoluteAxis")

Dim line2D13 As Line2D
Set line2D13 = axis2D4.GetItem("HDirection")

line2D13.ReportName = 1

Dim line2D14 As Line2D
Set line2D14 = axis2D4.GetItem("VDirection")

line2D14.ReportName = 2

Dim reference68 As Reference
Set reference68 = part1.CreateReferenceFromBRepName("WireREdge:(Wire:(Brp:(Sketch.1;22);None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MonoFond;MFBRepVersion_CXR15)", sketch1)

Dim geometricElements9 As GeometricElements
Set geometricElements9 = factory2D4.CreateIntersections(reference68)

Dim geometry2D5 As Geometry2D
Set geometry2D5 = geometricElements9.Item("Mark.1")

geometry2D5.Construction = True

Dim circle2D4 As Circle2D
Set circle2D4 = factory2D4.CreateClosedCircle(0#, 2.830222, 0.4)

circle2D4.ReportName = 3

Dim constraints4 As Constraints
Set constraints4 = sketch4.Constraints

Dim reference69 As Reference
Set reference69 = part1.CreateReferenceFromObject(circle2D4)

Dim constraint28 As Constraint
Set constraint28 = constraints4.AddMonoEltCst(catCstTypeRadius, reference69)

constraint28.Mode = catCstModeDrivingDimension

sketch4.CloseEdition

part1.InWorkObject = hybridBody1

part1.Update

part1.InWorkObject = body1

Dim reference70 As Reference
Set reference70 = part1.CreateReferenceFromBRepName("WireREdge:(Wire:(Brp:(Sketch.1;16);None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

Dim hybridShapeExtract4 As HybridShapeExtract
Set hybridShapeExtract4 = hybridShapeFactory1.AddNewExtract(reference70)

hybridShapeExtract4.PropagationType = 3

hybridShapeExtract4.ComplementaryExtract = False

hybridShapeExtract4.IsFederated = False

hybridBody1.AppendHybridShape hybridShapeExtract4

part1.InWorkObject = hybridShapeExtract4

part1.Update

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim reference71 As Reference
Set reference71 = part1.CreateReferenceFromObject(sketch2)

Dim reference72 As Reference
Set reference72 = part1.CreateReferenceFromObject(hybridShapeExtract4)

Dim rib1 As Rib
Set rib1 = shapeFactory1.AddNewRibFromRef(reference71, reference72)

part1.Update

Dim reference73 As Reference
Set reference73 = part1.CreateReferenceFromBRepName("WireREdge:(Wire:(Brp:(Sketch.1;14);None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

Dim hybridShapeExtract5 As HybridShapeExtract
Set hybridShapeExtract5 = hybridShapeFactory1.AddNewExtract(reference73)

hybridShapeExtract5.PropagationType = 3

hybridShapeExtract5.ComplementaryExtract = False

hybridShapeExtract5.IsFederated = False

hybridBody1.AppendHybridShape hybridShapeExtract5

part1.InWorkObject = hybridShapeExtract5

part1.Update

Dim reference74 As Reference
Set reference74 = part1.CreateReferenceFromObject(sketch3)

Dim reference75 As Reference
Set reference75 = part1.CreateReferenceFromObject(hybridShapeExtract5)

Dim rib2 As Rib
Set rib2 = shapeFactory1.AddNewRibFromRef(reference74, reference75)

part1.Update

Dim reference76 As Reference
Set reference76 = part1.CreateReferenceFromBRepName("WireREdge:(Wire:(Brp:(Sketch.1;22);None:(Limits1:();Limits2:());Cf11:());WithPermanentBody;WithoutBuildError;WithSelectingFeatureSupport;MFBRepVersion_CXR15)", sketch1)

Dim hybridShapeExtract6 As HybridShapeExtract
Set hybridShapeExtract6 = hybridShapeFactory1.AddNewExtract(reference76)

hybridShapeExtract6.PropagationType = 3

hybridShapeExtract6.ComplementaryExtract = False

hybridShapeExtract6.IsFederated = False

hybridBody1.AppendHybridShape hybridShapeExtract6

part1.InWorkObject = hybridShapeExtract6

part1.Update

Dim reference77 As Reference
Set reference77 = part1.CreateReferenceFromObject(sketch4)

Dim reference78 As Reference
Set reference78 = part1.CreateReferenceFromObject(hybridShapeExtract6)

Dim rib3 As Rib
Set rib3 = shapeFactory1.AddNewRibFromRef(reference77, reference78)

part1.Update

End Sub

