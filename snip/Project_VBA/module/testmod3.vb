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
Set reference1 = part1.CreateReferenceFromName("Selection_RSur:(Face:(Brp:(Pad.2;2);None:();Cf11:());Pad.2_ResultOUT;Z0;G4252)")

Dim sketch1 As Sketch
Set sketch1 = sketches1.Add(reference1)

Dim arrayOfVariantOfDouble1(8)
arrayOfVariantOfDouble1(0) = 20#
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

Dim point2D1 As Point2D
Set point2D1 = factory2D1.CreatePoint(0#, 20#)

point2D1.ReportName = 3

Dim line2D3 As Line2D
Set line2D3 = factory2D1.CreateLine(0#, 0#, 0#, 20#)

line2D3.ReportName = 4

Dim point2D2 As Point2D
Set point2D2 = axis2D1.GetItem("Origin")

line2D3.StartPoint = point2D2

line2D3.EndPoint = point2D1

Dim constraints1 As Constraints
Set constraints1 = sketch1.Constraints

Dim reference2 As Reference
Set reference2 = part1.CreateReferenceFromObject(line2D3)

Dim reference3 As Reference
Set reference3 = part1.CreateReferenceFromObject(line2D2)

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddBiEltCst(catCstTypeVerticality, reference2, reference3)

constraint1.Mode = catCstModeDrivingDimension

Dim point2D3 As Point2D
Set point2D3 = factory2D1.CreatePoint(-10#, 30#)

point2D3.ReportName = 5

Dim line2D4 As Line2D
Set line2D4 = factory2D1.CreateLine(0#, 20#, -10#, 30#)

line2D4.ReportName = 6

line2D4.StartPoint = point2D1

line2D4.EndPoint = point2D3

Dim point2D4 As Point2D
Set point2D4 = factory2D1.CreatePoint(-10#, 40#)

point2D4.ReportName = 7

Dim line2D5 As Line2D
Set line2D5 = factory2D1.CreateLine(-10#, 30#, -10#, 40#)

line2D5.ReportName = 8

line2D5.StartPoint = point2D3

line2D5.EndPoint = point2D4

Dim reference4 As Reference
Set reference4 = part1.CreateReferenceFromObject(line2D5)

Dim reference5 As Reference
Set reference5 = part1.CreateReferenceFromObject(line2D2)

Dim constraint2 As Constraint
Set constraint2 = constraints1.AddBiEltCst(catCstTypeVerticality, reference4, reference5)

constraint2.Mode = catCstModeDrivingDimension

sketch1.CloseEdition

part1.InWorkObject = body1

part1.Update

End Sub

