Sub CATMain()

Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim partDocument1 As PartDocument
Set partDocument1 = documents1.Add("Part")

Dim part1 As Part
Set part1 = partDocument1.Part

Dim hybridBodies1 As HybridBodies
Set hybridBodies1 = part1.HybridBodies

Dim hybridBody1 As HybridBody
Set hybridBody1 = hybridBodies1.Item("Geometrical Set.1")

Dim sketches1 As Sketches
Set sketches1 = hybridBody1.HybridSketches

Dim sketch1 As Sketch
Set sketch1 = sketches1.Item("Sketch.1")

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
Set circle2D1 = factory2D1.CreateClosedCircle(0#, 0#, 0#)

Dim point2D1 As Point2D
Set point2D1 = axis2D1.GetItem("Origin")

circle2D1.CenterPoint = point2D1

circle2D1.ReportName = 3

circle2D1.Construction = True

Dim constraints1 As Constraints
Set constraints1 = sketch1.Constraints

Dim reference1 As Reference
Set reference1 = part1.CreateReferenceFromObject(circle2D1)

Dim constraint1 As Constraint
Set constraint1 = constraints1.AddMonoEltCst(catCstTypeRadius, reference1)

constraint1.Mode = catCstModeDrivingDimension

sketch1.CloseEdition

part1.InWorkObject = hybridBody1

part1.UpdateObject sketch1

Dim parameters1 As Parameters
Set parameters1 = part1.Parameters

Dim length1 As Length
Set length1 = parameters1.Item("diameter1")

length1.Value = 50#

part1.Update

part1.InWorkObject = sketch1

Set factory2D1 = sketch1.OpenEdition()

sketch1.CloseEdition

part1.InWorkObject = hybridBody1

part1.UpdateObject sketch1

part1.InWorkObject = sketch1

Set factory2D1 = sketch1.OpenEdition()

End Sub

