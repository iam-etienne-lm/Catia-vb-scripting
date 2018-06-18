Option Explicit
Sub getnames55()
'v001

Dim specsAndGeomWindow1 As SpecsAndGeomWindow
Set specsAndGeomWindow1 = CATIA.ActiveWindow

Dim productDocument1 As ProductDocument
Set productDocument1 = CATIA.ActiveDocument

Dim product1 As Product
Set product1 = productDocument1.Product

Dim products1 As Products
Set products1 = product1.Products


Debug.Print products1.Count
Debug.Print CATIA.ActiveDocument.Product.Products.Count
Debug.Print CATIA.ActiveDocument.Count

End Sub
