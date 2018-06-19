Option Explicit
Sub getnames55()
On Error Resume Next
'v010 list names and changes the 10th level1 one

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

Dim intPartcount As Integer
intPartcount = CATIA.ActiveDocument.Product.Products.Count
'intPartcount = 10

Dim strNewname As String
strNewname = "changedname1" & intPartcount

Dim i As Integer
For i = 1 To intPartcount
    Debug.Print CATIA.ActiveDocument.Product.Products.Item(i).Name
    If i = 10 Then
        CATIA.ActiveDocument.Product.Products.Item(i).Name = strNewname
    End If
    
Next i





'Debug.Print CATIA.ActiveDocument.Product.Products.Item(1).Count

'Dim product8 As Product
'Set product8 = CATIA.ActiveDocument.Product.Products.AddNewComponent("Part", "part00108")



End Sub
