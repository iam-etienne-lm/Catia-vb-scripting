Sub test3parts()

Dim documents1 As Documents
Set documents1 = CATIA.Documents

Dim productDocument1 As ProductDocument
Set productDocument1 = documents1.Add("Product")


Dim product1 As Product
Set product1 = productDocument1.Product

Dim products1 As Products
Set products1 = product1.Products
'LEVEL1
Dim product2 As Product
Set product2 = products1.AddNewComponent("Part", "part00101")

Dim product3 As Product
Set product3 = products1.AddNewComponent("Part", "part00102")

Dim product4 As Product
Set product4 = products1.AddNewComponent("Part", "part00103")

End sub
