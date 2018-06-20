Sub catmain()
Dim prod1 As ProductDocument
Set prod1 = CATIA.ActiveDocument
Set prodInd = prod1.Product

Dim moveArray(11)
moveArray(0) = 1
moveArray(1) = 0
moveArray(2) = 0
moveArray(3) = 0
moveArray(4) = 1
moveArray(5) = 0
moveArray(6) = 0
moveArray(7) = 0
moveArray(8) = 1
moveArray(9) = -20 'translates component 1000 x direction of parent's coordinate space
moveArray(10) = 0
moveArray(11) = 0
prodInd.Move.Apply moveArray

End Sub
