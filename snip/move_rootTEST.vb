Sub catmain()
Dim proddoc1 As ProductDocument
Set proddoc1 = CATIA.ActiveDocument
Set prod1 = proddoc1.Product 'Subtility resides here. catia.activedomument can't be moved while catia.activedocument.product can be!
      'indeed a document is not the same as a product.
      
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
moveArray(9) = -20 
moveArray(10) = 0
moveArray(11) = 0
prod1.Move.Apply moveArray

End Sub
