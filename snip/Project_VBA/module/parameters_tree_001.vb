Sub CATMain()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As Part
Set part1 = partDocument1.Part

Dim parameters1 As Parameters
Set parameters1 = part1.Parameters

Dim angle1 As Angle
Set angle1 = parameters1.CreateDimension("", "ANGLE", 0#)

angle1.Rename "Angle Branches"

angle1.Value = 40#

Dim parameters2 As Parameters
Set parameters2 = part1.Parameters

Dim length1 As Length
Set length1 = parameters2.CreateDimension("", "LENGTH", 0#)

length1.Rename "Rayon Perçage"

length1.Value = 10#

Dim parameters3 As Parameters
Set parameters3 = part1.Parameters

Dim length2 As Length
Set length2 = parameters3.CreateDimension("", "LENGTH", 0#)

length2.Rename "Pas de Branches"

length2.Value = 1#

Dim parameters4 As Parameters
Set parameters4 = part1.Parameters

Dim angle2 As Angle
Set angle2 = parameters4.CreateDimension("", "ANGLE", 0#)

angle2.Rename "Angle début matière supportée"

angle2.Value = 40#

Dim parameters5 As Parameters
Set parameters5 = part1.Parameters

Dim length3 As Length
Set length3 = parameters5.CreateDimension("", "LENGTH", 0#)

length3.Rename "Rayon branches"

length3.Value = 0.4

End Sub

