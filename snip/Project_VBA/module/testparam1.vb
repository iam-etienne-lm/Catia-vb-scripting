Sub CATMain()

Dim partDocument1 As PartDocument
Set partDocument1 = CATIA.ActiveDocument

Dim part1 As Part
Set part1 = partDocument1.Part

Dim parameters1 As Parameters
Set parameters1 = part1.Parameters

Dim realParam1 As RealParam
Set realParam1 = parameters1.CreateReal("", 0#)

realParam1.Rename "Angle branches"

realParam1.Value = 40#

part1.Update

Set partDocument1 = CATIA.ActiveDocument

Dim selection1 As Selection
Set selection1 = partDocument1.Selection

selection1.Clear

selection1.Add realParam1

selection1.Delete

End Sub

