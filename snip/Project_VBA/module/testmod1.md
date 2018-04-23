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

Dim sketch1 As Sketch
Set sketch1 = sketches1.Item("Sketch.2")

part1.InWorkObject = sketch1

Dim factory2D1 As Factory2D
Set factory2D1 = sketch1.OpenEdition()

Dim circle2D1 As Circle2D
Set circle2D1 = factory2D1.CreateClosedCircle(0#, 0#, 22.36068)

Dim geometricElements1 As GeometricElements
Set geometricElements1 = sketch1.GeometricElements

Dim axis2D1 As Axis2D
Set axis2D1 = geometricElements1.Item("AbsoluteAxis")

Dim point2D1 As Point2D
Set point2D1 = axis2D1.GetItem("Origin")

circle2D1.CenterPoint = point2D1

circle2D1.ReportName = 2

sketch1.CloseEdition

Dim shapes1 As Shapes
Set shapes1 = body1.Shapes

Dim pad1 As Pad
Set pad1 = shapes1.Item("Pad.1")

part1.InWorkObject = pad1

part1.Update

Dim shapeFactory1 As ShapeFactory
Set shapeFactory1 = part1.ShapeFactory

Dim pad2 As Pad
Set pad2 = shapeFactory1.AddNewPad(sketch1, 20#)

part1.Update

End Sub

