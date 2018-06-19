Sub create-ScratchProduct()

Dim oProductDoc As Document
Set oProductDoc = CATIA.Documents.Add("Product")
'name to be changed

Dim specsAndGeomWindow1 As SpecsAndGeomWindow
Set specsAndGeomWindow1 = CATIA.ActiveWindow

specsAndGeomWindow1.WindowState = catWindowStateMaximized

Dim productDocument1 As ProductDocument
Set productDocument1 = CATIA.ActiveDocument

Dim product1 As Product
Set product1 = productDocument1.Product

Dim products1 As Products
Set products1 = product1.Products

Dim product2 As Product
Set product2 = products1.AddNewComponent("Part", "part00103")

End Sub
