Sub CATMain()

Dim productDocument1 As ProductDocument
Set productDocument1 = CATIA.ActiveDocument

Dim product1 As Product
Set product1 = productDocument1.Product

Dim products1 As Products
Set products1 = product1.Products

        'KNOWN UNRESOLVED ERROR ADDNEWCOMPONENT
Dim product2 As Product
On Error Resume Next
Set product2 = products1.AddNewComponent("Part", "y002")
MsgBox Err.Description
On Error GoTo 0


End Sub
